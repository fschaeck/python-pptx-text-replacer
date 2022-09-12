# python-pptx-text-replacer #

A Python script using python-pptx to replace text in a PowerPoint presentation (pptx) in:
  - all shapes with text frames (also inside grouped shapes)
  - chart categories
  - table cells
while preserving the original character's formatting.

The script is work in progress!

There might still be quirks in PowerPoint-presentations this script doesn't deal with well. Please open an issue stating your problem and attach (if in any way possible) the appropriate pptx file for me to be able to re-produce your problem. If you can't include the pptx itself, maybe you can post the part of the script's output where it goes wrong.

## Installation ##

Use pip to install the package python-pptx-text-replacer from PyPI:

```
python -m pip install python-pptx-text-replacer
```

This will also install python-pptx if it isn't installed already and
create the command (wrapper) python-pptx-text-replacer.

Thereafter you can use the package on the command line or use the class
python_pptx_text_replacer.TextReplacer in your own Python modules.

## Usage on command line  ##

You can execute the script in two ways:
1. Using the command wrapper `/...whatever-path-pip-created-it-under.../python-pptx-text-replacer`
2. Using the module itself: `python -m python_pptx_text_replacer.TextReplacer`

The following is, what you get, if you start the script with the parameter --help:

```
usage: TextReplacer.py [-h] --match <match> --replace <replacement> [--verbose] [--quiet] [--regex] --input <input file> --output <output file> [--slides <list of slide numbers to process>] [--text-frames] [--no-text-frames]
                       [--tables] [--no-tables] [--charts] [--no-charts]

This package implements text replacement in Powerpoint files in pptx format.

The text is searched and replaced in all possible places while preserving the
original character's formatting.

Text replacement can be configured to leave certain slides untouched (by specifying
which slides should be processed), or to not touching text in tables, charts or
text frames in any of the shapes.

This package can be imported and the class python_pptx_text_replacer used directly
or it can be called as main and given parameters to define what needs to be done.

options:
  -h, --help            show this help message and exit
  --match <match>, -m <match>
                        the string to look for and to be replaced
  --replace <replacement>, -r <replacement>
                        the replacement for all the matches' occurrences
  --verbose, -v         print detailed structure of and changes made in presentation file
  --quiet, -q           don't even print the changes that are done
  --regex, -x           use match strings as regular expressions
  --input <input file>, -i <input file>
                        the file to replace the text in
  --output <output file>, -o <output file>
                        the file to write the changed presentation to
  --slides <list of slide numbers to process>, -s <list of slide numbers to process>
                        A comma-separated list of slide numbers (1-based) to restrict processing to, i.e. '2,4,6-10'
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

=================================================================
python-pptx-text-replacer v0.0.5post0 (c) Frank Schäckermann 2022
```

### Examples using the command line  ###

Let's assume you have the script python_pptx_text_replacer.py in your current directory, Python is installed correctly and the python-pptx module is installed as well.
Let's further assume you have a PowerPoint presentation file original.pptx in that same directory.

#### Replacing the string 'FY2021' with 'FY2122' in the whole presentation ####

Use the command

```
python -m python_pptx_text_replacer.TextReplacer -m FY2021 -r FY2122 -i ./original.pptx -o ./changed.pptx
```

and will find the changed presentation in the same directory in the file changed.pptx.

#### Replacing the strings 'FY2021' and 'FY1918' with 'FY2122' and 'FY2019' respectively in one run everywhere in the presentation ####
Use the command

```
python -m python_pptx_text_replacer.TextReplacer -m FY2021 -r FY2122 -m FY1918 -r FY2019 -i ./original.pptx -o ./changed.pptx
```

#### Replacing the string 'FY2021' with 'FY2122' but only in all chart categories ####
Use the command

```
python -m python_pptx_text_replacer.TextReplacer -m FY2021 -r FY2122 --no-tables --no-text-frames -i ./original.pptx -o ./changed.pptx
```

#### Replacing the string 'FY2021' with 'FY2122' but only in table headers and cells ####
Use the command

```
python -m python_pptx_text_replacer.TextReplacer -m FY2021 -r FY2122 --no-charts --no-text-frames -i ./original.pptx -o ./changed.pptx
```

#### Replacing the string 'FY2021' with 'FY2122' but only in all the shapes' text frames ####
Use the command

```
python -m python_pptx_text_replacer.TextReplacer -m FY2021 -r FY2122 --no-charts --no-tables -i ./original.pptx -o ./changed.pptx
```

#### Replacing the string 'FY2021' with 'FY2122' everywhere except on the 4th and 6th slide ####
Use the command

```
python -m python_pptx_text_replacer.TextReplacer -m FY2021 -r FY2122 --slides '1-3,5,7-' -i ./original.pptx -o ./changed.pptx
```

### Examples using the module in your own Python program ###

You need to import the module with

```
from python_pptx_text_replacer import TextReplacer
```
and then use the class TextReplacer as shown in below examples.

The parameters to the constructor of TextReplacer are:
1. (positional, required): name of file with presentation to process.
2. tables (named, optional): if True (default), tables will be processed, if False, tables will be ignored. 
3. charts (named, optional): if True (default), charts will be processed, if False, charts will be ignored. 
4. textframes (named, optional): if True (default), textframes will be processed, if False, textframes will be ignored. 
5. slides (named, optional): comma separated list of slide numbers to process. If not specified, all slides will be processed.
6. verbose (named, optional): Default is False. Will be used as default for each call to TextReplacer.replace_text.
6. quiet (named, optional): Default is False. Will be used as default for each call to TextReplacer.replace_text.

The parameter to the function TextReplacer.replace_text is a list of tuples of the form ( match,replacement ).
All match/replace-actions are done in the sequence the tuples appear in the list.
To avoid unforeseen results, the function is doing some sanity-checks on the list of replacements and prints a warning if it finds anything that might lead to unintended results.
Make sure, you understand what is going on, if you ignore these warnings!

There are two optional parameters to the function replace_text:
1. verbose
2. quiet
Their defaults are, what was specified when TextReplacer was created or False, if those where not specified on creation of TextReplacer. 

With verbose=True the function will print a detailed structure of the presentation and all the changes it is doing.
With verbose=False only the changes will be listed and finally with quiet=True not even those changes will be printed.
In any case - if there are any warnings or errors, they will be printed at the end - even with quiet=True.

The function replace_text can be called multiple times with different match/replace tuples. But be aware, that the sanity-checks will only include the current replacement tupels and won't look at former ones!

The presentation can be saved as often as you wish in between calls to replace_text() by using the function write_presentation_to_file.

#### Replacing the string 'FY2021' with 'FY2122' in the whole presentation ####

```
from python_pptx_text_replacer import TextReplacer
replacer = TextReplacer("original.pptx", slides='',
                        tables=True, charts=True, textframes=True)
replacer.replace_text( [ ('FY2021','FY2122') ] )
replacer.write_presentation_to_file("./changed.pptx")
```

#### Replacing the strings 'FY2021' and 'FY1918' with 'FY2122' and 'FY2019' respectively in one run everywhere in the presentation ####

```
from python_pptx_text_replacer import TextReplacer
replacer = TextReplacer("original.pptx", slides='',
                        tables=True, charts=True, textframes=True)
replacer.replace_text( [ ('FY2021','FY2122'),('FY1918','FY2019') ] )
replacer.write_presentation_to_file("./changed.pptx")
```

#### Replacing the string 'FY2021' with 'FY2122' but only in all chart categories ####

```
from python_pptx_text_replacer import TextReplacer
replacer = TextReplacer("original.pptx", slides='',
                        tables=False, charts=True, textframes=False)
replacer.replace_text( [ ('FY2021','FY2122') ] )
replacer.write_presentation_to_file("./changed.pptx")
```

#### Replacing the string 'FY2021' with 'FY2122' but only in table headers and cells ####

```
from python_pptx_text_replacer import TextReplacer
replacer = TextReplacer("original.pptx", slides='',
                        tables=True, charts=False, textframes=False)
replacer.replace_text( [ ('FY2021','FY2122') ] )
replacer.write_presentation_to_file("./changed.pptx")
```

#### Replacing the string 'FY2021' with 'FY2122' but only in all the shapes' text frames ####

```
from python_pptx_text_replacer import TextReplacer
replacer = TextReplacer("original.pptx", slides='',
                        tables=False, charts=False, textframes=True)
replacer.replace_text( [ ('FY2021','FY2122') ] )
replacer.write_presentation_to_file("./changed.pptx")
```

#### Replacing the string 'FY2021' with 'FY2122' everywhere except on the 4th and 6th slide ####

```
from python_pptx_text_replacer import TextReplacer
replacer = TextReplacer("original.pptx", slides='1-3,5,7-',
                        tables=True, charts=True, textframes=True)
replacer.replace_text( [ ('FY2021','FY2122') ] )
replacer.write_presentation_to_file("./changed.pptx")
```
