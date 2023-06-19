#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Add an appearance stream, to a file."""

# get pdf-annotate from
# https://github.com/plangrid/pdf-annotate
# and add to path
import sys
sys.path.append('/some/dir/pdf-annotate')

# https://github.com/plangrid/pdf-annotate
from pdf_annotate import PdfAnnotator, Location, Appearance

color = (1, 0, 0)
transparency = 0.2

a = PdfAnnotator('empty.pdf')
a.add_annotation(
    'square',
    Location(x1=50, y1=700, x2=100, y2=750, page=0),
    Appearance(stroke_color=color,
               stroke_width=0,
               stroke_transparency=transparency,
               fill=color,
               fill_transparency=transparency),
)
a.write('file_with_appearance_stream.pdf')
