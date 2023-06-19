#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Convert hyperlinks into attachments.

Only hyperlinks to local files are processed, remote hyperlinks (e.g. http:,
https:, ftp:, mailto:) are ignored.

Author: Marcelo Andrioni
https://github.com/marceloandrioni

"""

import os
import sys
import re
from pathlib import Path
import warnings
import argparse
from pikepdf import (Pdf, AttachedFileSpec, Name, Dictionary, Permissions,
                     Encryption)
from gooey import Gooey, GooeyParser


# Use CLI (instead of GUI) if the CLI arguments were passed.
# https://github.com/chriskiehl/Gooey/issues/449#issuecomment-534056010
if len(sys.argv) > 1:
    if '--ignore-gooey' not in sys.argv:
        sys.argv.append('--ignore-gooey')


def attach_file(pdf, annot, filespec):

    pushpin = Dictionary(Type=Name('/Annot'),
                         Subtype=Name('/FileAttachment'),
                         Name=Name('/PushPin'),
                         FS=filespec.obj,
                         Rect=annot['/Rect'],
                         Contents=filespec.description,   # file description
                         C=(1.0, 1.0, 0.0),   # color
                         T=None,   # author
                         M=None)   # modification date, e.g.: 'D:20210101000000'

    # Get an appearance stream from another file and use it in place of the
    # PushPin icon.
    #
    # Note: as each viewer (Adobe, Evince, Okular, etc) has it own
    # implementation of the default icons (Graph, PushPin, Paperclip, Tag), the
    # icons do not have a standard appearance. Also, some viewers (e.g. the
    # ones in Google Chrome and Microsoft Edge), only follow the v2 pdf 
    # standard, where there is no default icons, so no icon is show.
    # The fix is to use an appearance stream, basically a dictionary listing
    # exactly how the annotation/icon should be represented. This way the 
    # annotation/icon shows the same in all viewers.
    # I could not create an appearance stream from scratch, so the solution was
    # to use pdf-annotate (https://github.com/plangrid/pdf-annotate) to save an
    # appearance stream to a file and reuse it here.

    as_file = (Path(__file__).parent
               / 'appearance_stream/file_with_appearance_stream.pdf')
    with Pdf.open(as_file) as pdf2:
        pushpin['/AP'] = pdf.copy_foreign(pdf2.pages[0]['/Annots'][0])['/AP']
    
    return pdf.make_indirect(pushpin)


def user_args():

    description = 'Convert hyperlinks into attachments.'
    parser = GooeyParser(description=description,
                         allow_abbrev=False)

    parser.add_argument('infile',
                        type=lambda x: Path(x),
                        help='Input pdf file with hyperlinks to local files.',
                        widget='FileChooser',
                        gooey_options={
                            'wildcard': 'PDF file (*.pdf)|*.pdf',
                            'message': 'Select input pdf file'})

    parser.add_argument('outfile',
                        type=lambda x: Path(x),
                        help='Output pdf file with attachments.',
                        widget='FileSaver',
                        gooey_options={
                            'wildcard': 'PDF file (*.pdf)|*.pdf',
                            'message': 'Select output pdf file'})

    args = parser.parse_args()

    if [args.infile.suffix, args.outfile.suffix] != ['.pdf', '.pdf']:
        raise argparse.ArgumentTypeError('Input/Output files must be pdf files.')

    if not args.infile.exists():
        raise argparse.ArgumentTypeError(f"Input file '{args.infile}' does "
                                         "not exist.")

    if args.outfile.exists() and os.path.samefile(args.infile, args.outfile):
        raise argparse.ArgumentTypeError("Input/Output files can't be the same.")

    return args


def check_uri(uri, relative_path):
    """Path if uri is a valid local file, else, None."""

    if uri is None:
        return

    uri = str(uri)

    # do nothing if remote link
    if uri.startswith(('http:', 'https:', 'ftp:', 'mailto:')):
        return

    # if uri starts with "file:" the path is absolute, else, is relative
    if uri.startswith('file:///'):

        # three slashes:
        # file:///C:/eclipse/eclipse.ini
        # file://///mynetwordir/eclipse.ini
        uri = Path(re.sub('^file:///', '', uri))

    else:

        # path is relative to the input pdf file
        uri = relative_path / uri

    if not uri.exists():
        raise ValueError(f"Local file '{uri}' does not exist.")

    return uri


@Gooey(required_cols=1,
       progress_regex=r"^Page (?P<current>\d+)/(?P<total>\d+)$",
       progress_expr="current / total * 100")
def main():

    args = user_args()

    print(f'Input file: {args.infile}')

    pdf = Pdf.open(args.infile)

    # Main loop based on:
    # Post: https://stackoverflow.com/a/65977239
    # Author: https://stackoverflow.com/users/14282700/shivang-raj
    filespecs = {}
    for page in pdf.pages:

        print(f'Page {page.index + 1}/{len(pdf.pages)}')

        for idx, annot in enumerate(page.get('/Annots', {})):

            uri = annot.get('/A', {}).get('/URI')

            uri = check_uri(uri, args.infile.parent)

            if uri is None:
                continue

            print(f"  Attaching local file '{uri}'")

            # avoid attaching copies of the same file if two or more hyperlinks
            # reference the file.
            if uri not in filespecs:
                filespecs[uri] = AttachedFileSpec.from_filepath(
                    pdf,
                    uri,
                    description=uri.name)

            # replace the hyperlink annotation with the attached file annotation
            page['/Annots'][idx] = attach_file(pdf, annot, filespecs[uri])

    fnames = [x.name for x in filespecs.keys()]
    if len(fnames) != len(set(fnames)):
        warnings.warn(
            'There is hyperlinks referencing files with the same name '
            '(not taking into account the directory path). The hyperlinks will '
            'reference the correct attached files in the output pdf, however, '
            'the lateral attachment bar in pdf viewers (e.g. Adobe Acrobat '
            'Reader, Firefox) will only display the first of the homonymous files.'
        )

    # Default page layout when opening the pdf file. Some viewers may ignore it.
    # https://pikepdf.readthedocs.io/en/latest/topics/pagelayout.html
    pdf.Root.PageLayout = Name.OneColumn
    pdf.Root.PageMode = Name.UseOutlines

    # Do not allow a regular user to modify the file.
    # This is simply a protection so that the user does not accidentally remove
    # the attached file annotation when viewing the file in Adobe Acrobat Reader.
    allow = Permissions(accessibility=True,
                        extract=True,
                        modify_annotation=False,
                        modify_assembly=False,
                        modify_form=False,
                        modify_other=False,
                        print_lowres=True,
                        print_highres=True)
    encryption = Encryption(user='', owner='admin123', allow=allow)

    print(f'Output file: {args.outfile}')
    # linearize=True: Enables creating linear or "fast web view", where the
    # file's contents are organized sequentially so that a viewer can begin
    # rendering before it has the whole file. As a drawback, it tends to make
    # files larger.
    # https://pikepdf.readthedocs.io/en/latest/api/main.html#pikepdf.Pdf.save
    pdf.save(args.outfile, linearize=True, encryption=encryption)

    pdf.close()

    print('Done!')


if __name__ == '__main__':
    main()
