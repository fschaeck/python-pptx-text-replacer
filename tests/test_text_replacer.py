import os
import sys
import unittest

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

    def do_test(self,textframes,tables,charts,replacements,expected_stdout,expected_stderr):
        pass

    def test_01_change_nothing(self):
        self.do_test(False,False,False,[('cell','CELL')],[''],None)
