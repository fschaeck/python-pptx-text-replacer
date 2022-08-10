# python-pptx-text-replacer #

A Python script using python-pptx to replace text in a PowerPoint presentation (pptx) in:
  - all shapes with text frames (also inside grouped shapes)
  - chart categories
  - table cells
while preserving the original character's formatting.

The script is work in progress!

There might still be quirks in PowerPoint-presentations this script doesn't deal with well. Please open an issue stating your problem and attach (if in any way possible) the appropriate pptx file for me to be able to re-produce your problem. If you can't include the pptx itself, maybe you can post the part of the script's output where it goes wrong.

The following is, what you get, if you start the script with the parameter --help:

```
usage: python-pptx-text-replacer [-h] --match <match> --replace <replacement>
                                 --input <input file> --output <output file>
                                 [--slides <list of slide numbers to process>]
                                 [--text-frames] [--no-text-frames] [--tables]
                                 [--no-tables] [--charts] [--no-charts]

This module implements text replacement in Powerpoint files in pptx format.

The text is searched and replaced in all possible places while preserving the
original character's formatting.

Text replacement can be configured to leave certain slides untouched (by specifying
which slides should be processed), or to not touching text in tables, charts or
text frames in any of the shapes.

This module can be imported and the class python_pptx_text_replacer used directly
or it can be called as main and given parameters to define what needs to be done.

optional arguments:
  -h, --help            show this help message and exit
  --match <match>, -m <match>
                        the string to look for and to be replaced
  --replace <replacement>, -r <replacement>
                        the replacement for all the matches' occurrences
  --input <input file>, -i <input file>
                        the file to replace the text in
  --output <output file>, -o <output file>
                        the file to write the changed presentation to
  --slides <list of slide numbers to process>, -s <list of slide numbers to process>
                        A comma-separated list of slide numbers (1-based) to
                        restrict processing to, i.e. '2,4,6-10'
  --text-frames, -f     process text frames in any shape as well (default)
  --no-text-frames, -F  do not process any text frames in shapes
  --tables, -t          process tables as well (default)
  --no-tables, -T       do not process tables and their cells
  --charts, -c          process chart categories as well (default)
  --no-charts, -C       do not process charts and their categories

The parameters --match and --replace can be specified multiple times.
They are paired up in the order of their appearance.

The slide list given with --slides must be a comma-separated list of
slide numbers from 1 to the number of slides contained in the presentation
or slide number ranges of the kind '4-16'. If the second number is omitted,
like in '4-' the range includes everything from the slide identified by the
first number up to the last slide in the file.
```

## Examples using the script as main ##

Let's assume you have the script python_pptx_text_replacer.py in your current directory, Python is installed correctly and the python-pptx module is installed as well.
Let's further assume you have a PowerPoint presentation file original.pptx in that same directory.

### Replacing the string 'FY2021' with 'FY2122' in the whole presentation ###

You execute the command

```
python ./python_pptx_text_replacer.py -m FY2021 -r FY2122 -i ./original.pptx -o ./changed.pptx
```

and will find the changed presentation in the same directory in the file changed.pptx.

### Replacing the strings 'FY2021' and 'FY1918' with 'FY2122' and 'FY2019' respectively in one run everywhere in the presentation ###
Use the command

```
python ./python_pptx_text_replacer.py -m FY2021 -r FY2122 -m FY1918 -r FY2019 -i ./original.pptx -o ./changed.pptx
```

### Replacing the string 'FY2021' with 'FY2122' but only in all chart categories ###
Use the command

```
python ./python_pptx_text_replacer.py -m FY2021 -r FY2122 --no-tables --no-text-frames -i ./original.pptx -o ./changed.pptx
```

### Replacing the string 'FY2021' with 'FY2122' but only in table headers and cells ###
Use the command

```
python ./python_pptx_text_replacer.py -m FY2021 -r FY2122 --no-charts --no-text-frames -i ./original.pptx -o ./changed.pptx
```

### Replacing the string 'FY2021' with 'FY2122' but only in all the shapes' text frames ###
Use the command

```
python ./python_pptx_text_replacer.py -m FY2021 -r FY2122 --no-charts --no-tables -i ./original.pptx -o ./changed.pptx
```

### Replacing the string 'FY2021' with 'FY2122' everywhere except on the 4th and 6th slide ###
Use the command

```
python ./python_pptx_text_replacer.py -m FY2021 -r FY2122 --slides '1-3,5,7-' -i ./original.pptx -o ./changed.pptx
```

## Examples using the module in your own Python program ##

You need to import the module with

```
from python_pptx_text_replacer import python_pptx_text_replacer
```
and then use the class python_pptx_text_replacer in the following specified ways...

Note: The parameter to the function python_pptx_text_replacer.replace_text is a list of tuples of the form ( match,replacement ). All match/replace-actions are done in the sequence the tuples appear in the list. Also this function can be called multiple times with different match/replace tuples and the presentation can be saved in between if need be by repeated calls to the function write_presentation_file().

### Replacing the string 'FY2021' with 'FY2122' in the whole presentation ###

```
from python_pptx_text_replacer import python_pptx_text_replacer
replacer = python_pptx_text_replacer("original.pptx", slides='',
                                     tables=True, charts=True, textframes=True)
replacer.replace_text( [ ('FY2021','FY2122') ] )
replacer.write_presentation_to_file("./changed.pptx")
```

### Replacing the strings 'FY2021' and 'FY1918' with 'FY2122' and 'FY2019' respectively in one run everywhere in the presentation ###

```
from python_pptx_text_replacer import python_pptx_text_replacer
replacer = python_pptx_text_replacer("original.pptx", slides='',
                                     tables=True, charts=True, textframes=True)
replacer.replace_text( [ ('FY2021','FY2122'),('FY1918','FY2019') ] )
replacer.write_presentation_to_file("./changed.pptx")
```

### Replacing the string 'FY2021' with 'FY2122' but only in all chart categories ###

```
from python_pptx_text_replacer import python_pptx_text_replacer
replacer = python_pptx_text_replacer("original.pptx", slides='',
                                     tables=False, charts=True, textframes=False)
replacer.replace_text( [ ('FY2021','FY2122') ] )
replacer.write_presentation_to_file("./changed.pptx")
```

### Replacing the string 'FY2021' with 'FY2122' but only in table headers and cells ###

```
from python_pptx_text_replacer import python_pptx_text_replacer
replacer = python_pptx_text_replacer("original.pptx", slides='',
                                     tables=True, charts=False, textframes=False)
replacer.replace_text( [ ('FY2021','FY2122') ] )
replacer.write_presentation_to_file("./changed.pptx")
```

### Replacing the string 'FY2021' with 'FY2122' but only in all the shapes' text frames ###

```
from python_pptx_text_replacer import python_pptx_text_replacer
replacer = python_pptx_text_replacer("original.pptx", slides='',
                                     tables=False, charts=False, textframes=True)
replacer.replace_text( [ ('FY2021','FY2122') ] )
replacer.write_presentation_to_file("./changed.pptx")
```

### Replacing the string 'FY2021' with 'FY2122' everywhere except on the 4th and 6th slide ###

```
from python_pptx_text_replacer import python_pptx_text_replacer
replacer = python_pptx_text_replacer("original.pptx", slides='1-3,5,7-',
                                     tables=True, charts=True, textframes=True)
replacer.replace_text( [ ('FY2021','FY2122') ] )
replacer.write_presentation_to_file("./changed.pptx")
```
