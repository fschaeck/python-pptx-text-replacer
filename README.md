# python-pptx-text-replacer

A Python script using python-pptx to replace text in a PowerPoint presentation (pptx) in:
  - all shapes with text frames (also inside grouped shapes)
  - chart categories
  - table cells

The script is work in progress!

There might still be quirks in PowerPoint-presentations this script doesn't deal with well. Please open an issue stating your problem and (if in any way possible) the appropriate pptx file for me to be able to re-produce your problem. If you can't include the pptx itself, maybe you can post the part of the script's output where it goes wrong.

```
usage: python-pptx-text-replacer.py [-h] --match <match> --replace
                                    <replacement> --input <input file>
                                    --output <output file> [--tables]
                                    [--no-tables] [--charts] [--no-charts]
                                    [--slides <list of slide numbers>]

This module implements text replacement in Powerpoint files in pptx format.
The text is searched and replaced in all possible places.

optional arguments:
  -h, --help            show this help message and exit
  --match <match>, -m <match>
                        the match to look for
  --replace <replacement>, -r <replacement>
                        the matches' replacement
  --input <input file>, -i <input file>
                        the file to replace the text in
  --output <output file>, -o <output file>
                        the file to write the changed presentation to
  --tables, -t          process tables as well (default)
  --no-tables, -T       do not process tables and their cells
  --charts, -c          process chart categories as well (default)
  --no-charts, -C       do not process charts and their categories
  --slides <list of slide numbers>, -s <list of slide numbers>
                        A comma-separated list of slide numbers (1-based) to
                        restrict processing to, i.e. '2,4,6-10'

The parameters --match and --replace can be specified multiple times. They are
paired up in the order of their appearance.
```
