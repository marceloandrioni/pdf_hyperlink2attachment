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
from collections import namedtuple
import warnings
import argparse
import tempfile
from pikepdf import (Pdf, AttachedFileSpec, Name, Dictionary, Permissions,
                     Encryption)
from gooey import Gooey, GooeyParser

indir = Path(__file__).resolve().parent / 'external/pdf_annotate'
sys.path.append(str(indir))
from pdf_annotate import PdfAnnotator, Location, Appearance


# Use CLI (instead of GUI) if the CLI arguments were passed.
# https://github.com/chriskiehl/Gooey/issues/449#issuecomment-534056010
if len(sys.argv) > 1:
    if '--ignore-gooey' not in sys.argv:
        sys.argv.append('--ignore-gooey')


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


class Hyperlinks2Attachments:

    def __init__(self, pdf):
        self._pdf = pdf

    @property
    def pdf(self):
        return self._pdf

    def _get_local_hyperlinks(self):
        """Loop pages/annotations and return only those where hyperlinks point
        to local files."""

        HyperLink = namedtuple('HyperLink', ['page_idx',
                                             'annotation_idx',
                                             'position',
                                             'uri'])

        hlinks = []
        for page in self.pdf.pages:

            if page.get('/Annots') is None:
                continue

            for annot_idx, annot in enumerate(page['/Annots']):

                if annot.get('/A') is None:
                    continue

                if annot['/A'].get('/URI') is None:
                    continue

                uri = str(annot['/A'].get('/URI'))

                # skip if remote file
                if uri.startswith(('http:', 'https:', 'ftp:', 'mailto:')):
                    continue

                # check if local file has absolute or relative path
                if uri.startswith('file:///'):

                    # path is abolute
                    # file:///C:/eclipse/eclipse.ini   (local file in local disk)
                    # file://///mynetwordir/eclipse.ini   (local file in network disk)
                    uri = Path(re.sub('^file:///', '', uri))

                else:

                    # path is relative to the pdf file
                    uri = Path(self.pdf.filename).resolve().parent / uri

                if not uri.exists():
                    raise ValueError(f"Local file '{uri}' does not exist.")

                hlinks.append(HyperLink(page.index,
                                        annot_idx,
                                        annot['/Rect'],
                                        uri))

        return hlinks

    def _create_appearance_stream(self, rect):
        """Return an appearance stream."""

        # Note: As each viewer (Adobe, Evince, Okular, etc) has it own
        # implementation of the default icons (Graph, PushPin, Paperclip, Tag), the
        # icons do not have a standard appearance. Also, some viewers (e.g. the
        # ones in Google Chrome and Microsoft Edge), only accept the v2 pdf
        # standard, where there is no default icons, so no icon is show.
        # The fix is to use an appearance stream, basically a dictionary listing
        # exactly how the annotation/icon should be represented. This way the
        # annotation/icon shows the same in all viewers.
        # I could not create an appearance stream from scratch, so the solution was
        # to use pdf-annotate (https://github.com/plangrid/pdf-annotate) to save an
        # appearance stream to a file and reuse it.

        # temporary files
        tmp_file = tempfile.mktemp(suffix='.pdf')
        tmp_file2 = tempfile.mktemp(suffix='.pdf')

        # create empty file tmp_file
        with Pdf.new() as fp:
            fp.add_blank_page()
            fp.save(tmp_file)

        # add annotation to empty file and save it as tmp_file2
        fp = PdfAnnotator(tmp_file)
        fp.add_annotation(
            'square',
            Location(page=0,
                     x1=float(rect[0]),
                     y1=float(rect[1]),
                     x2=float(rect[2]),
                     y2=float(rect[3])),
            Appearance(stroke_color=(1, 0, 0),
                       stroke_width=1,
                       stroke_transparency=0.5))
        fp.write(tmp_file2)

        # copy the appearance stream from tmp_file2 to the main pdf
        with Pdf.open(tmp_file2) as pdf2:
            ap = self.pdf.copy_foreign(pdf2.pages[0]['/Annots'][0])['/AP']

        # remove temporary files
        os.remove(tmp_file)
        os.remove(tmp_file2)

        return ap

    def _attach_file(self, position, filespec):
        """Attach local file to pdf."""

        pushpin = Dictionary(Type=Name('/Annot'),
                             Subtype=Name('/FileAttachment'),
                             Name=Name('/PushPin'),
                             FS=filespec.obj,
                             Rect=position,
                             Contents=filespec.description,   # file description
                             C=(1.0, 1.0, 0.0),   # color
                             T=None,   # author
                             M=None)   # modification date, e.g.: 'D:20210101000000'

        # Get an appearance stream and use it instead of the PushPin icon.
        # Unlike the default (Graph, PushPin, Paperclip, Tag), the appearance
        # stream has the same looks in all viewers
        pushpin['/AP'] = self._create_appearance_stream(position)

        return self.pdf.make_indirect(pushpin)

    def _warn_if_same_name(self, files):
        if len(files) != len(set(files)):
            warnings.warn(
                'There is hyperlinks referencing files with the same name '
                '(not taking into account the directory path). The hyperlinks '
                'will reference the correct attached files in the output pdf, '
                'however, the lateral attachment bar in some pdf viewers (e.g.: '
                'Firefox) will only display the first of the homonymous files.'
            )

    def hyperlinks2attachments(self):

        hlinks = self._get_local_hyperlinks()
        if len(hlinks) == 0:
            print('No local hyperlink to attach.')
            return

        filespecs = {}
        for idx, hlink in enumerate(hlinks, start=1):

            print(f'Attaching {idx}/{len(hlinks)}')
            print(f"  Local file '{hlink.uri}'")

            # avoid attaching copies of the same file if two or more
            # hyperlinks reference the file.
            if hlink.uri not in filespecs:
                filespecs[hlink.uri] = AttachedFileSpec.from_filepath(
                    self.pdf,
                    hlink.uri,
                    description=hlink.uri.name)

            # replace the hyperlink annotation with the attached file annotation
            self.pdf.pages[hlink.page_idx]['/Annots'][hlink.annotation_idx] = \
                self._attach_file(hlink.position,
                                  filespecs[hlink.uri])

        self._warn_if_same_name([x.name for x in filespecs])

    def save(self, outfile):

        # Default page layout when opening the pdf file. Some viewers may ignore it.
        # https://pikepdf.readthedocs.io/en/latest/topics/pagelayout.html
        self.pdf.Root.PageLayout = Name.OneColumn
        self.pdf.Root.PageMode = Name.UseOutlines

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

        # linearize=True: Enables creating linear or "fast web view", where the
        # file's contents are organized sequentially so that a viewer can begin
        # rendering before it has the whole file. As a drawback, it tends to make
        # files larger.
        # https://pikepdf.readthedocs.io/en/latest/api/main.html#pikepdf.Pdf.save
        self.pdf.save(outfile, linearize=True, encryption=encryption)


def hyperlinks2attachments(infile, outfile):

    infile = Path(infile)
    outfile = Path(outfile)

    if not infile.exists():
        raise ValueError(f'Infile {infile} must exist.')

    if infile.suffix != '.pdf':
        raise ValueError(f'Infile {infile} must be a pdf file.')

    if outfile.suffix != '.pdf':
        raise ValueError(f'Outfile {outfile} must be a pdf file.')

    with Pdf.open(infile) as pdf:
        h2a = Hyperlinks2Attachments(pdf)
        h2a.hyperlinks2attachments()
        h2a.save(outfile)


@Gooey(required_cols=1,
       progress_regex=r"^Attaching (?P<current>\d+)/(?P<total>\d+)$",
       progress_expr="current / total * 100")
def main():

    args = user_args()

    print(f'Input file: {args.infile}')
    print(f'Output file: {args.outfile}')

    hyperlinks2attachments(args.infile, args.outfile)

    print('Done!')


if __name__ == '__main__':
    main()
