# -*- coding: utf-8 -*-

from __future__ import print_function, unicode_literals

import os
import sys
import unittest
from io import open as open, StringIO
from python_pptx_text_replacer import TextReplacer

ENCODING = 'utf-8'

PY2 = sys.version_info[0] == 2
if PY2:
    def make_unicode(strg, encoding):
        if type(strg) == str:
            return unicode(strg, encoding)
        else:
            return strg
else:
    class unicode(object):  # @ReservedAssignment
        pass
    def make_unicode(strg, encoding):
        if type(strg) == bytes:
            return strg.decode(encoding)
        else:
            return strg
    def unichr(char):  # @ReservedAssignment
        return chr(char)


class Capture(object):
    def __init__(self, stdin_data):
        if stdin_data is not None:
            if PY2 and type(stdin_data) != unicode or not PY2 and type(stdin_data) != str:
                raise ValueError('Programming error: Capture(stdin_data) not unicode.')
        self._stdin_data = stdin_data  # must be unicode
        self._stdin = None
        self._stdout_data = []
        self._stderr_data = []

    def __enter__(self):
        self._stdout = sys.stdout
        sys.stdout = self._stringio_out = StringIO()
        self._stderr = sys.stderr
        sys.stderr = self._stringio_err = StringIO()
        if self._stdin_data is not None:
            self._stdin = sys.stdin
            sys.stdin = self._stringio_in = StringIO(self._stdin_data)
        else:
            self._stdin = None
        return self

    def __exit__(self, *args):  # @UnusedVariable
        content = self._stringio_out.getvalue()
        if len(content) == 0:
            pass
        elif content.endswith('\n'):
            self._stdout_data.extend(lne+'\n' for lne in content[:-1].split('\n'))
        else:
            self._stdout_data.extend(lne+'\n' for lne in content.split('\n'))
        sys.stdout = self._stdout
        del self._stringio_out

        content = self._stringio_err.getvalue()
        if len(content) == 0:
            pass
        elif content.endswith('\n'):
            self._stderr_data.extend(lne+'\n' for lne in content[:-1].split('\n'))
        else:
            self._stderr_data.extend(lne+'\n' for lne in content.split('\n'))
        sys.stderr = self._stderr
        del self._stringio_err

        if self._stdin is not None:
            del self._stringio_in
            sys.stdin = self._stdin

    def stdout(self):
        return self._stdout_data

    def stderr(self):
        return self._stderr_data


class test_text_replacer(unittest.TestCase):

    def setUp(self):
        pass


    def tearDown(self):
        pass


    def make_list(self, encoding, content):
        if type(content) == str or PY2 and type(content) == unicode:
            if len(content) == 0:
                return []
            content = make_unicode(content, encoding)
            if content.endswith('\n'):
                return list(lne+'\n' for lne in content[:-1].split('\n'))
            else:
                return list(lne+'\n' for lne in content.split('\n'))
        elif type(content) == list:
            return list(make_unicode(lne, encoding) for lne in content)
        raise ValueError('Programming error: invalid content parameter type ({}) for make_list'
                         .format(type(content)))


    def check_output(self, encoding, content_name, expected_content, actual_content):
        MISSING_MARKER = '<missing>'
        UNEXPECTED_MARKER = '<unexpected>'

        list1 = self.make_list(encoding, expected_content)
        list2 = self.make_list(encoding, actual_content)

        content_name = os.path.basename(content_name)
        tag1 = 'expected '+content_name
        tag2 = 'actual '+content_name

        max_lst_len = max(len(list1), len(list2))
        if max_lst_len == 0:
            return []

        # make sure both lists have same length
        list1.extend([None] * (max_lst_len - len(list1)))
        list2.extend([None] * (max_lst_len - len(list2)))

        max_txt_len_1 = max(list(len(UNEXPECTED_MARKER)
                                 if txt is None
                                 else 3*len(txt)-2*len(txt.rstrip('\r\n'))
                                 for txt in list1)+[len(tag1)])
        max_txt_len_2 = max(list(len(MISSING_MARKER)
                                 if txt is None
                                 else 3*len(txt)-2*len(txt.rstrip('\r\n'))
                                 for txt in list2)+[len(tag2)])

        diff = ['']
        equal = True
        diff.append('|  No | ? | {tag1:<{txtlen1}.{txtlen1}s} | {tag2:<{txtlen2}.{txtlen2}s} |'
                    .format(tag1=tag1, tag2=tag2, txtlen1=max_txt_len_1, txtlen2=max_txt_len_2))
        for i, (x, y) in enumerate(zip(list1, list2)):
            if x != y:
                equal = False
                if x is not None and y is not None and x.rstrip('\r\n') == y.rstrip('\r\n'):
                    x = x.replace('\n', '\\n').replace('\r', '\\r')
                    y = y.replace('\n', '\\n').replace('\r', '\\r')
            diff.append('| {idx:>3d} | {equal:1.1s} | {line1:<{txtlen1}.{txtlen1}s} | {line2:<{txtlen2}.{txtlen2}s} |'  # noqa: E501
                        .format(idx=i+1,
                                equal=(' ' if x == y else '*'),
                                txtlen1=max_txt_len_1,
                                txtlen2=max_txt_len_2,
                                line1=UNEXPECTED_MARKER
                                      if x is None
                                      else x.rstrip('\r\n'),  # .replace(' ', '\N{MIDDLE DOT}'),
                                line2=MISSING_MARKER
                                      if y is None
                                      else y.rstrip('\r\n')))  # .replace(' ', '\N{MIDDLE DOT}')))

        return [] if equal else diff

    def do_test(self,input_file,
                     textframes,tables,charts,
                     slides,
                     replacements,
                     expected_stdout,
                     expected_stderr):
        rc = 0
        with Capture(None) as capture:
            try:
                replacer = TextReplacer(input_file,
                                        textframes=textframes,
                                        tables=tables,
                                        charts=charts,
                                        slides=slides)
                replacer.replace_text(replacements)
            except ValueError as err:
                print(str(err),file=sys.stderr)
                rc = 1
        result = []
        if expected_stdout is not None:
             result.extend(self.check_output(ENCODING,'stdout',expected_stdout,capture.stdout()))
        if expected_stderr is not None:
            result.extend(self.check_output(ENCODING,'stderr',expected_stderr,capture.stderr()))

        if len(result) > 0:
            try:
                result = '\n'.join(result)
                self.fail(result)
            except TypeError:
                self.fail(str(result))


    def test_01_change_nothing(self):
        self.do_test('tests/data/Test-Presentation.pptx',False,False,False,'',[('cell','CELL')],
"""Presentation[tests/data/Test-Presentation.pptx]
  Slide[1, id=256] with title 'Trying a table'
    Shape[0, id=2, type=PLACEHOLDER (14)]
      ... skipped
    Shape[1, id=4, type=TABLE (19)]
      Table[4,4]
        ... skipped
  Slide[2, id=257] with title 'A Chart'
    Shape[0, id=2, type=PLACEHOLDER (14)]
      ... skipped
    Shape[1, id=3, type=CHART (3)]
      Chart of type COLUMN_STACKED (52)
        ... skipped
  Slide[3, id=258] with title 'A Textbox'
    Shape[0, id=2, type=PLACEHOLDER (14)]
      ... skipped
    Shape[1, id=3, type=TEXT_BOX (17)]
      ... skipped
  Slide[4, id=259] with title 'Grouped Shapes'
    Shape[0, id=2, type=PLACEHOLDER (14)]
      ... skipped
    Shape[1, id=5, type=GROUP (6)]
      Shape[0, id=3, type=AUTO_SHAPE (1)]
        ... skipped
      Shape[1, id=4, type=AUTO_SHAPE (1)]
        ... skipped
"""
,'')
