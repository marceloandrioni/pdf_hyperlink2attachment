#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
In a pdf file, convert hyperlinks into attachments.

Only hyperlinks to local files are processed, remote hyperlinks (e.g. http://,
https://, ftp://, mailto:) are ignored.

Author: Marcelo Andrioni
https://github.com/marceloandrioni

"""

import os
from pathlib import Path
import argparse
from pikepdf import Pdf, AttachedFileSpec, Name, Dictionary


def attach_file(pdf, page, annot, uri):

    filespec = AttachedFileSpec.from_filepath(
        pdf,
        uri,
        description=uri.name)

    # https://www.pdftron.com/api/PDFTronSDK/dotnet/pdftron.PDF.Annots.FileAttachment.html
    # Note that FileAttachment icons can differ in their appearance dimensions,
    # so you may want to match these Rectangle dimensions or the aspect ratio
    # to avoid a squished or stretched appearance :
    # Graph : 40 x 40
    # PushPin : 28 x 40
    # Paperclip : 14 x 34
    # Tag : 40 x 32

    # using a random name (not one of the four allowed ones) to draw a
    # transparent block over the hyperlinked text
    icon_name = 'None'

    pushpin = Dictionary(
        Type=Name('/Annot'),
        Subtype=Name('/FileAttachment'),
        Name=Name(f'/{icon_name}'),
        FS=filespec.obj,
        Rect=annot['/Rect'],
        Contents=uri.name,   # file description
        C=(1.0, 1.0, 0.0),   # color
        T=None,   # author
        M=None,   # modification date, e.g.: 'D:20210101000000'
    )

    return pdf.make_indirect(pushpin)


def argfile(x):

    x = Path(x)
    if x.suffix != '.pdf':
        raise argparse.ArgumentTypeError('Input/Output files must be pdf files.')

    return x


def cli_args():

    description = 'In a pdf file, convert hyperlinks into attachments.'
    parser = argparse.ArgumentParser(description=description)

    parser.add_argument('infile',
                        type=lambda x: argfile(x),
                        help='Input pdf file with hyperlinks to local files.')

    parser.add_argument('outfile',
                        type=lambda x: argfile(x),
                        help='Output pdf file with attachments.')

    parser.add_argument('-O', '--overwrite',
                        action='store_true',
                        default=False,
                        help='Overwrite outfile if exists (Default: False).')

    args = parser.parse_args()

    if not args.infile.exists():
        raise argparse.ArgumentTypeError(f"Input file '{args.infile}' does "
                                         "not exist.")

    if os.path.samefile(args.infile, args.outfile):
        raise argparse.ArgumentTypeError("Input/Output files can't be the same.")

    if args.outfile.exists() and not args.overwrite:
        raise argparse.ArgumentTypeError(f"Output file {args.outfile} exist. "
                                         "Use -O flag to overwrite.")

    return args


def main():

    args = cli_args()

    print(f'Input file: {args.infile}')

    pdf = Pdf.open(args.infile)

    # Main loop based on:
    # Post: https://stackoverflow.com/a/65977239
    # Author: https://stackoverflow.com/users/14282700/shivang-raj
    for page in pdf.pages:
        for idx, annot in enumerate(page.get('/Annots', {})):
            uri = annot.get('/A', {}).get('/URI')

            if uri is None:
                continue

            uri = str(uri)

            # do nothing if remote link
            if uri.startswith(('http://', 'https://', 'ftp://', 'mailto:')):
                continue

            uri = args.infile.parent / uri

            print(f"  Attaching local file: '{uri}'")

            if not uri.exists():
                raise ValueError(f"Local file '{uri}' does not exist.")

            # replace the hyperlink annotation with the attached file annotation
            page['/Annots'][idx] = attach_file(pdf, page, annot, uri)

    # linearize=True: Enables creating linear or "fast web view", where the
    # file's contents are organized sequentially so that a viewer can begin
    # rendering before it has the whole file. As a drawback, it tends to make
    # files larger.
    pdf.save(args.outfile, linearize=True)
    pdf.close()

    print(f'Output file: {args.outfile}')

    print('Done!')


if __name__ == '__main__':
    main()
