"""
parse PDF invoice file and extract data to Excel
"""

import getopt
import os
import re
import sys
import pickle
from collections import defaultdict as Dict
from itertools import chain
import logging

import fitz
import pandas as pd
import pdfplumber as pb

logging.basicConfig(level=logging.ERROR,
                    filename= __name__ + '.log',
                    datefmt='%Y-%m-%d %H:%M:%S',
                    format='%(asctime)s - %(levelname)s - %(funcName)s - %(processName)s - %(threadName)s - %(message)s')
logger = logging.getLogger(__name__ + '_logger')


class Extractor(object):
    def __init__(self, path):
        self.file = path

    def _check_file(self):
        if not isinstance(self.file, str) or not os.path.isfile(self.file) or not self.file.endswith(('.pdf', '.PDF')):
            return {'error': 'not a valid pdf file.'}
        return True

    def _load_data(self):
        try:
            doc = fitz.open(self.file)
            page = doc.loadPage(0)
            words = page.getTextWords()

            words = [{'x0': int(round(word[0])), 'y0': int(round(word[1])), 'x1': int(round(word[2])),
                      'y1': int(round(word[3])), 'word': word[4]} for word in words]
            words = sorted(words, key=lambda v: v['x0'])
            words = sorted(words, key=lambda v: v['y0'])
            maxY = max(w['y1']for w in words)
            delta = 30
            for idx, word in enumerate(words):
                words[idx]['y0'] = maxY + delta - word['y0']
                words[idx]['y1'] = maxY + delta - word['y1']

            pdf = pb.open(self.file)
            page = pdf.pages[0]
            words2 = page.extract_words()
            words2 = [{'x0': int(round(word['x0'])), 'y0': int(round(word['top'])), 'x1': int(round(word['x1'])),
                       'y1': int(round(word['bottom'])), 'word': word['text']} for word in words2]
            words2 = sorted(words2, key=lambda v: v['x0'])
            words2 = sorted(words2, key=lambda v: v['y0'])

            lines = [{'x0': round(line['x0']),
                      'y0': round(line['y0']),
                      'x1': round(line['x1']),
                      'y1': round(line['y1']),
                      'width': round(line['width']),
                      'height': round(line['height'])} for line in page.lines]
            lines = sorted(lines, key=lambda v: v['x0'])
            lines = sorted(lines, key=lambda v: v['y0'])
        except Exception as e:
            return {'error': e}

        return {'words': words, 'words2': words2, 'lines': lines}

    @staticmethod
    def _find_nearest_val(vals, val):
        delta = [abs(v-val) for v in vals]
        idx = delta.index(min(delta))
        return vals[idx]

    def _fill_line(self, lines):
        hlines = [line for line in lines if line['width'] > 0]  # ????????????
        hlines = sorted(hlines, key=lambda h: h['width'], reverse=True)[:-2]  # ?????????????????????
        vlines = [line for line in lines if line['height'] > 0]  # ????????????

        # ??????????????????????????????
        ys = [line['y0'] for line in chain(hlines, vlines)] + [line['y1'] for line in chain(hlines, vlines)]
        xs = [line['x0'] for line in chain(hlines, vlines)] + [line['x1'] for line in chain(hlines, vlines)]
        for idx, line in enumerate(hlines):
            for k in ['x0', 'y0', 'x1', 'y1']:
                series = xs.copy() if 'x' in k else ys.copy()
                series.remove(line[k])
                hlines[idx][k] = self._find_nearest_val(series, line[k])
        for idx, line in enumerate(vlines):
            for k in ['x0', 'y0', 'x1', 'y1']:
                series = xs.copy() if 'x' in k else ys.copy()
                series.remove(line[k])
                vlines[idx][k] = self._find_nearest_val(series, line[k])
        # ??????????????????
        maxX = max(int(line['x1']) for line in chain(hlines, vlines))
        minX = min(int(line['x0']) for line in chain(hlines, vlines))
        minY = min(int(line['y0']) for line in chain(hlines, vlines))
        maxY = max(int(line['y1']) for line in chain(hlines, vlines))

        thline = {'x0': minX, 'y0': minY, 'x1': maxX, 'y1': minY}  # ????????????
        bhline = {'x0': minX, 'y0': maxY, 'x1': maxX, 'y1': maxY}  # ????????????
        lvline = {'x0': minX, 'y0': minY, 'x1': minX, 'y1': maxY}  # ????????????
        rvline = {'x0': maxX, 'y0': minY, 'x1': maxX, 'y1': maxY}  # ????????????

        hlines.insert(0, thline)
        hlines.append(bhline)
        vlines.insert(0, lvline)
        vlines.append(rvline)
        return hlines, vlines

    @staticmethod
    def _is_point_in_rect(point, rect):
        """???????????????????????????"""
        px, py = point
        p1, p2, p3, p4 = rect.values()
        if p1[0] <= px <= p2[0] and p1[1] <= py <= p3[1]:
            return True
        else:
            return False

    @staticmethod
    def _find_cross_points(hlines, vlines):
        points = []
        delta = 1
        for vline in vlines:
            vx0 = vline['x0']
            vy0 = vline['y0']
            vy1 = vline['y1']
            for hline in hlines:
                hx0 = hline['x0']
                hy0 = hline['y0']
                hx1 = hline['x1']
                if (hx0 - delta) <= vx0 <= (hx1 + delta) and (vy0 - delta) <= hy0 <= (vy1 + delta):
                    points.append((int(vx0), int(hy0)))
        return points

    @staticmethod
    def _find_rects(cross_points):
        # ????????????
        X = sorted(set([int(p[0]) for p in cross_points]))
        Y = sorted(set([int(p[1]) for p in cross_points]))
        df = pd.DataFrame(index=Y, columns=X)
        for p in cross_points:
            x, y = int(p[0]), int(p[1])
            df.loc[y, x] = 1
        df = df.fillna(0)
        # ????????????
        rects = []
        COLS = len(df.columns) - 1
        ROWS = len(df.index) - 1
        for row in range(ROWS):
            for col in range(COLS):
                p0 = df.iat[row, col]  # ?????????????????????????????????
                cnt = col + 1
                while cnt <= COLS:
                    p1 = df.iat[row, cnt]
                    p2 = df.iat[row + 1, col]
                    p3 = df.iat[row + 1, cnt]
                    if p0 and p1 and p2 and p3:
                        rects.append({'p0': (df.columns[col], df.index[row]),
                                      'p1': (df.columns[cnt], df.index[row]),
                                      'p2': (df.columns[cnt], df.index[row+1]),
                                      'p3': (df.columns[col], df.index[row+1])})
                        break
                    else:
                        cnt += 1
        return rects

    def _name_rects(self, rects):
        rects = sorted(rects, key=lambda r: r['p0'][0])
        rects = sorted(rects, key=lambda r: r['p0'][1], reverse=True)
        return {f'r{idx+1}': rect for idx, rect in enumerate(rects)}

    def _put_words_into_rect(self, words, rects):
        # ???words?????????????????????????????????
        groups = {'IN': Dict(list), 'OUT': Dict(list)}
        for name, r in rects.items():
            groups['IN'][name] = []
        for word in words:
            p = ((word['x0'] + word['x1']) // 2, (word['y0'] + word['y1']) // 2)
            is_word_put_into_group = False
            for name, r in rects.items():
                if self._is_point_in_rect(p, r):
                    is_word_put_into_group = True
                    groups['IN'][name].append(word)
                    break

            if not is_word_put_into_group:
                groups['OUT'][word['x0']].append(word)
        return groups

    @staticmethod
    def _find_text_by_same_line(group, delta=1):
        words = {}
        group = sorted(group, key=lambda x: x['x0'])
        for w in group:
            bottom = int(w['bottom'])
            text = w['text']
            k1 = [bottom - i for i in range(delta)]
            k2 = [bottom + i for i in range(delta)]
            k = set(k1 + k2)
            flag = False
            for kk in k:
                if kk in words:
                    words[kk] = words.get(kk, '') + text
                    flag = True
                    break
            if not flag:
                words[bottom] = words.get(bottom, '') + text
        return words

    def _split_words_into_diff_line(self, groups):
        groups2 = {}
        for k, g in groups.items():
            words = self._find_text_by_same_line(g, 3)
            groups2[k] = words
        return groups2

    @staticmethod
    def _index_of_y(x, rects):
        for index, r in enumerate(rects):
            if x == r[2][0][0]:
                return index + 1 if index + 1 < len(rects) else None
        return None

    @staticmethod
    def _find_outer(words):
        df = pd.DataFrame()
        for pos, text in words.items():
            if re.search(r'??????$', text):  # ????????????
                df.loc[0, '????????????'] = text
            elif re.search(r'????????????', text):  # ????????????
                num = ''.join(re.findall(r'[0-9]+', text))
                df.loc[0, '????????????'] = num
            elif re.search(r'????????????', text):  # ????????????
                num = ''.join(re.findall(r'[0-9]+', text))
                df.loc[0, '????????????'] = num
            elif re.search(r'????????????', text):  # ????????????
                date = ''.join(re.findall(
                    r'[0-9]{4}???[0-9]{1,2}???[0-9]{1,2}???', text))
                df.loc[0, '????????????'] = date
            elif '????????????' in text and '?????????' in text:  # ?????????
                text1 = re.search(r'?????????:\d+', text)[0]
                num = ''.join(re.findall(r'[0-9]+', text1))
                df.loc[0, '?????????'] = num
                text2 = re.search(r'????????????:\d+', text)[0]
                num = ''.join(re.findall(r'[0-9]+', text2))
                df.loc[0, '????????????'] = num
            elif '????????????' in text:
                num = ''.join(re.findall(r'[0-9]+', text))
                df.loc[0, '????????????'] = num
            elif '?????????' in text:
                num = ''.join(re.findall(r'[0-9]+', text))
                df.loc[0, '?????????'] = num
            elif re.search(r'?????????', text):
                items = re.split(r'?????????:|??????:|?????????:|?????????:', text)
                items = [item for item in items if re.sub(
                    r'\s+', '', item) != '']
                df.loc[0, '?????????'] = items[0] if items and len(items) > 0 else ''
                df.loc[0, '??????'] = items[1] if items and len(items) > 1 else ''
                df.loc[0, '?????????'] = items[2] if items and len(items) > 2 else ''
                df.loc[0, '?????????'] = items[3] if items and len(items) > 3 else ''
        return df

    @staticmethod
    def _find_and_sort_rect_in_same_line(y, groups):
        same_rects_k = [k for k, v in groups.items() if k[1] == y]
        return sorted(same_rects_k, key=lambda x: x[2][0][0])

    def _find_inner(self, k, words, groups, groups2, free_zone_flag=False):
        df = pd.DataFrame()
        sort_words = sorted(words.items(), key=lambda x: x[0])
        text = [word for k, word in sort_words]
        context = ''.join(text)
        if '?????????' in context or '?????????' in context:
            y = k[1]
            x = k[2][0][0]
            same_rects_k = self._find_and_sort_rect_in_same_line(y, groups)
            target_index = self._index_of_y(x, same_rects_k)
            target_k = same_rects_k[target_index]
            group_context = groups2[target_k]
            prefix = '?????????' if '?????????' in context else '?????????'
            for pos, text in group_context.items():
                if '??????' in text:
                    name = re.sub(r'??????:', '', text)
                    df.loc[0, prefix + '??????'] = name
                elif '??????????????????' in text:
                    tax_man_id = re.sub(r'??????????????????:', '', text)
                    df.loc[0, prefix + '??????????????????'] = tax_man_id
                elif '???????????????' in text:
                    addr = re.sub(r'???????????????:', '', text)
                    df.loc[0, prefix + '????????????'] = addr
                elif '??????????????????' in text:
                    account = re.sub(r'??????????????????:', '', text)
                    df.loc[0, prefix + '??????????????????'] = account
        elif '?????????' in context:
            y = k[1]
            x = k[2][0][0]
            same_rects_k = self._find_and_sort_rect_in_same_line(y, groups)
            target_index = self._index_of_y(x, same_rects_k)
            target_k = same_rects_k[target_index]
            words = groups2[target_k]
            context = [v for k, v in words.items()]
            context = ''.join(context)
            df.loc[0, '?????????'] = context
        elif '????????????' in context:
            y = k[1]
            x = k[2][0][0]
            same_rects_k = self._find_and_sort_rect_in_same_line(y, groups)
            target_index = self._index_of_y(x, same_rects_k)
            target_k = same_rects_k[target_index]
            group_words = groups2[target_k]
            group_context = ''.join([w for k, w in group_words.items()])
            items = re.split(r'[(???]??????[)???]', group_context)
            b = items[0] if items and len(items) > 0 else ''
            s = items[1] if items and len(items) > 1 else ''
            df.loc[0, '????????????(??????)'] = b
            df.loc[0, '????????????(??????)'] = s
        elif '??????' in context:
            y = k[1]
            x = k[2][0][0]
            same_rects_k = self._find_and_sort_rect_in_same_line(y, groups)
            target_index = self._index_of_y(x, same_rects_k)
            if target_index:
                target_k = same_rects_k[target_index]
                group_words = groups2[target_k]
                group_context = ''.join([w for k, w in group_words.items()])
                df.loc[0, '??????'] = group_context
            else:
                df.loc[0, '??????'] = ''
        else:
            if free_zone_flag:
                return df, free_zone_flag
            y = k[1]
            x = k[2][0][0]
            same_rects_k = self._find_and_sort_rect_in_same_line(y, groups)
            if len(same_rects_k) == 8:
                free_zone_flag = True
                for kk in same_rects_k:
                    words = groups2[kk]
                    words = sorted(words.items(), key=lambda x: x[0]) if words and len(
                        words) > 0 else None
                    key = words[0][1] if words and len(words) > 0 else None
                    val = [word[1] for word in words[1:]
                           ] if key and words and len(words) > 1 else ''
                    val = '\n'.join(val) if val else ''
                    if key:
                        df.loc[0, key] = val
        return df, free_zone_flag

    def _search_inner(self, inner_groups):
        s = pd.Series(dtype=object)
        if 'r2' in inner_groups:
            try:
                words_in_line = self._merge_words_by_line(inner_groups['r2'])
                for line, words in words_in_line.items():
                    words = sorted(words, key=lambda w: w['x0'])
                    word = ''.join(str(w['word']) for w in words)
                    vals = re.split(r'[:???]', word)
                    if len(vals) > 1:
                        key, val = vals[:2]
                    else:
                        key = vals[0]
                        val = 0
                    s[key+'(?????????)'] = val
            except Exception as e:
                logger.error(f'error in r2: {e}')
        if 'r4' in inner_groups:
            try:
                words_in_line = self._merge_words_by_line(inner_groups['r4'])
                text = ''
                for line, words in words_in_line.items():
                    words = sorted(words, key=lambda w: w['x0'])
                    word = ''.join(str(w['word']) for w in words)
                    text += word
                s['?????????'] = text
            except Exception as e:
                logger.error(f'error in r4: {e}')
        if 'r5' in inner_groups:
            try:
                words_in_line = self._merge_words_by_line(inner_groups['r5'])
                vals = []
                for line, words in words_in_line.items():
                    words = sorted(words, key=lambda w: w['x0'])
                    word = ''.join(str(w['word']) for w in words)
                    vals.append(word)
                if len(vals) > 2:
                    s[vals[0]] = '\n'.join(str(v) for v in vals[1:-1])
                elif len(vals) == 2:
                    s[vals[0]] = ''
                else:
                    logger.error(f'not enough val in r5: {vals}')
            except Exception as e:
                logger.error(f'error in r5: {e}')
        for r in ['r6', 'r7', 'r8', 'r9', 'r11']:
            if r in inner_groups:
                try:
                    words_in_line = self._merge_words_by_line(inner_groups[r])
                    vals = []
                    for line, words in words_in_line.items():
                        words = sorted(words, key=lambda w: w['x0'])
                        word = ''.join(str(w['word']) for w in words)
                        vals.append(word)
                    if len(vals) > 0:
                        s[vals[0]] = '\n'.join(str(v) for v in vals[1:])
                    elif len(vals) == 1:
                        s[vals[0]] = ''
                    else:
                        logger.error(f'not enough val in {r}: {vals}')
                except Exception as e:
                    logger.error(f'error in {r}: {e}')
        if 'r10' in inner_groups:
            try:
                words_in_line = self._merge_words_by_line(inner_groups['r10'])
                vals = []
                for line, words in words_in_line.items():
                    words = sorted(words, key=lambda w: w['x0'])
                    word = ''.join(str(w['word']) for w in words)
                    vals.append(word)
                if len(vals) > 2:
                    s[vals[0]] = '\n'.join(str(v) for v in vals[1:-1])
                    s['??????(??????)'] = vals[-1]
                elif len(vals) == 2:
                    s[vals[0]] = ''
                    s['??????(??????)'] = vals[-1]
                else:
                    logger.error(f'not enough val in r10: {vals}')
            except Exception as e:
                logger.error(f'error in r10: {e}')
        if 'r12' in inner_groups:
            try:
                words_in_line = self._merge_words_by_line(inner_groups['r12'])
                vals = []
                for line, words in words_in_line.items():
                    words = sorted(words, key=lambda w: w['x0'])
                    word = ''.join(str(w['word']) for w in words)
                    vals.append(word)
                if len(vals) > 2:
                    s[vals[0]] = '\n'.join(str(v) for v in vals[1:-1])
                    s['??????(??????'] = vals[-1]
                elif len(vals) == 2:
                    s[vals[0]] = ''
                    s['??????(??????)'] = vals[-1]
                else:
                    logger.error(f'not enough val in r12: {vals}')
            except Exception as e:
                logger.error(f'error in r12: {e}')
        if 'r14' in inner_groups:
            try:
                words_in_line = self._merge_words_by_line(inner_groups['r14'])
                for line, words in words_in_line.items():
                    words = sorted(words, key=lambda w: w['x0'])
                    word = ''.join(str(w['word']) for w in words)
                    vals = re.split(r'[(???]??????[???)]', word)
                    if len(vals) >= 2:
                        upper, lower = vals[:2]
                    else:
                        upper = vals[0]
                        lower = ''
                    s['????????????(??????)'] = upper
                    s['????????????(??????)'] = lower
            except Exception as e:
                logger.error(f'error in r14: {e}')
        if 'r16' in inner_groups:
            try:
                words_in_line = self._merge_words_by_line(inner_groups['r16'])
                for line, words in words_in_line.items():
                    words = sorted(words, key=lambda w: w['x0'])
                    word = ''.join(str(w['word']) for w in words)
                    vals = re.split(r'[:???]', word)
                    if len(vals) > 1:
                        key, val = vals[:2]
                    else:
                        key = vals[0]
                        val = 0
                    s[key+'(?????????)'] = val
            except Exception as e:
                logger.error(f'error in r16: {e}')
        if 'r18' in inner_groups:
            try:
                words_in_line = self._merge_words_by_line(inner_groups['r18'])
                for line, words in words_in_line.items():
                    words = sorted(words, key=lambda w: w['x0'])
                    word = ''.join(str(w['word']) for w in words)
                    s['??????'] = word
            except Exception as e:
                logger.error(f'error in r18: {e}')
        return s

    def _search_outer(self, outer_groups):
        s = pd.Series(dtype=object)
        words = [word for gwords in outer_groups.values() for word in gwords]
        words_in_line = self._merge_words_by_line(words)
        for row_num, words in words_in_line.items():
            words = sorted(words, key=lambda w: w['x0'])
            text = ''.join(str(w['word']) for w in words)
            if re.search(r'[\u4e00-\u9fa5]{3,20}??????', text):  # ????????????
                s['????????????'] = re.findall(r'[\u4e00-\u9fa5]{3,20}??????', text)[0]
            for key in ['????????????', '????????????', '?????????', '????????????']:
                if key in text:
                    sep = re.compile(key + r'[???:\s]')
                    rule = re.compile(key + r'[???:\s]' + r'\d+')
                    vals = re.findall(rule, text)
                    val = vals[0] if len(vals) > 0 else ''
                    val = re.sub(sep, '', val)
                    s[key] = val
            if re.search(r'????????????', text):  # ????????????
                date = ''.join(re.findall(r'\d{4}???\d{1,2}???\d{1,2}???', text))
                s['????????????'] = date
            if re.search(r'?????????', text):
                items = re.split(r'?????????:|??????:|?????????:|?????????:', text)
                items = [item for item in items if re.sub(r'\s+', '', item) != '']
                s['?????????'] = items[0] if items and len(items) > 0 else ''
                s['??????'] = items[1] if items and len(items) > 1 else ''
                s['?????????'] = items[2] if items and len(items) > 2 else ''
                s['?????????'] = items[3] if items and len(items) > 3 else ''
        return s

    @staticmethod
    def _merge_words_by_line(words, delta=2):
        words_in_line = Dict(list)
        for word in words:
            row_num = round((word['y0'] + word['y1'])/2)
            row_range = set([row_num - i for i in range(1, delta+1)] + [row_num + i for i in range(1, delta+1)])
            if len(row_range & set(words_in_line.keys())) > 0:
                row_num = list(row_range & set(words_in_line.keys()))[0]
            words_in_line[row_num].append(word)
        return words_in_line

    def extract(self):
        if self._check_file() is not True:
            return self._check_file()
        data = self._load_data()
        if 'error' in data:
            return data
        words = data['words']
        # words2 = data['words2']
        lines = data['lines']

        hlines, vlines = self._fill_line(lines)
        cross_points = self._find_cross_points(hlines, vlines)
        rects = self._find_rects(cross_points)
        if len(rects) < 18 and os.path.isfile('rects.pickle'):
            with open('rects.pickle', 'rb') as f:
                rects = pickle.load(f)
        if len(rects) < 18:
            return {'error': 'can\'t get rects.'}
        named_rects = self._name_rects(rects)

        words_groups = self._put_words_into_rect(words, named_rects)
        inner = self._search_inner(words_groups['IN'])
        outer = self._search_outer(words_groups['OUT'])
        res = pd.concat([inner, outer])
        return res


def load_files(directory):
    """load files"""
    if not os.path.isdir(directory):
        return []
    path_in_folder = Dict(list)
    for root, _, files in os.walk(directory):
        for file_ in files:
            path = os.path.join(root, file_)
            folder_name = re.split(r'/|\\', root)[-1]
            if os.path.isfile(path) and file_.endswith(('.pdf', '.PDF')):
                path_in_folder[folder_name].append(path)
    return path_in_folder


def test():
    import cv2
    import numpy as np
    import matplotlib.pyplot as plt

    path = 'example/test.pdf'
    extractor = Extractor(path)
    data = extractor._load_data()
    if 'error' in data:
        return data
    words = data['words']
    # words2 = data['words2']
    lines = data['lines']

    hlines, vlines = extractor._fill_line(lines)
    cross_points = extractor._find_cross_points(hlines, vlines)
    rects = extractor._find_rects(cross_points)
    if len(rects) < 18 and os.path.isfile('rects.pickle'):
        with open('rects.pickle', 'rb') as f:
            rects = pickle.load(f)
    if len(rects) < 18:
        return {'error': 'can\'t get rects.'}
    named_rects = extractor._name_rects(rects)
    words_groups = extractor._put_words_into_rect(words, named_rects)
    for name, words in sorted(words_groups['IN'].items(), key=lambda x: int(x[0].replace('r', ''))):
        words = ' '.join([str(w['word']) for w in words])
        print(name, ': ', words)
    for name, words in sorted(words_groups['OUT'].items(), key=lambda x: x[0]):
        words = ' '.join([str(w['word']) for w in words])
        print(name, ': ', words)

    minX = min(int(line['x0']) for line in hlines)
    maxX = max(int(line['x1']) for line in hlines)
    minY = min(int(line['y0']) for line in vlines)
    maxY = max(int(line['y1']) for line in vlines)

    delta = 40
    width = maxX + minX + delta
    height = maxY + minY + delta

    mat1 = np.zeros((height, width, 3))
    for line in hlines:
        p0 = (line['x0'], height - line['y0'])
        p1 = (line['x1'], height - line['y1'])
        cv2.line(mat1, p0, p1, (0, 255, 0), 2)
    for line in vlines:
        p0 = (line['x0'], height - line['y0'])
        p1 = (line['x1'], height - line['y1'])
        cv2.line(mat1, p0, p1, (0, 0, 255), 2)
    plt.figure()
    plt.title('hlines+vlines')
    plt.imshow(mat1)

    mat2 = np.zeros((height, width, 3))
    for idx, r in named_rects.items():
        x0 = r['p0'][0]
        y0 = height - r['p0'][1]
        x1 = r['p2'][0]
        y1 = height - r['p2'][1]
        cv2.rectangle(mat2, (x0, y0), (x1, y1), (0, 255, 0), 2)
        cv2.putText(mat2, idx, ((x0+x1)//2, (y0+y1)//2), cv2.FONT_HERSHEY_TRIPLEX, 0.6, (0, 255, 0), 2)
    plt.figure()
    plt.title('rects')
    plt.imshow(mat2)

    mat3 = np.zeros((height, width, 3))
    for name, words in words_groups['IN'].items():
        # mat = np.zeros((height, width, 3))
        x0 = named_rects[name]['p0'][0]
        y0 = height - named_rects[name]['p0'][1]
        x1 = named_rects[name]['p2'][0]
        y1 = height - named_rects[name]['p2'][1]
        # cv2.rectangle(mat, (x0, y0), (x1, y1), (0, 255, 0), 2)
        cv2.rectangle(mat3, (x0, y0), (x1, y1), (0, 255, 0), 2)

        for word in words:
            x0 = word['x0']
            x1 = word['x1']
            y0 = word['y0']
            y1 = word['y1']
            p = (round((x0 + x1) / 2), height - round((y0 + y1) / 2))
            # cv2.circle(mat, p, 2, (255, 0, 0), 2)
            cv2.circle(mat3, p, 2, (255, 0, 0), 2)
        # plt.figure()
        # plt.title(f'words in rects:{name}')
        # plt.imshow(mat)
    for row, words in words_groups['OUT'].items():
        for word in words:
            x0 = word['x0']
            x1 = word['x1']
            y0 = word['y0']
            y1 = word['y1']
            p = (round((x0 + x1) / 2), height - round((y0 + y1) / 2))
            cv2.circle(mat3, p, 2, (255, 0, 0), 2)

    plt.figure()
    plt.title(f'words in rects')
    plt.imshow(mat3)
    plt.show()


def main():
    IN_PATH = '/Users/bluzy/Documents/Reciepts'
    OUT_PATH = 'result.xlsx'
    DEBUG = True
    # parse params
    opts, args = getopt.getopt(sys.argv[1:], 'p:ts:', ['test', 'path=', 'save=', 'debug'])
    for opt, arg in opts:
        if opt in ['-p', '--path']:
            IN_PATH = arg
        elif opt in ['--test', '-t']:
            IN_PATH = 'example'
        elif opt in ['--save', '-s']:
            OUT_PATH = arg
        elif opt == '--debug':
            DEBUG = True

    if DEBUG:
        test()
        sys.exit(0)
    # run program
    print(
        f'run {"test" if IN_PATH == "example" else "extracting"} mode, load data from directory {IN_PATH}.\n{"*" * 50}')
    files_path = load_files(IN_PATH)
    num = [len(paths) for _, paths in files_path.items()]
    print(f'total {len(num)} folders, {sum(num)} file(s) to parse.\n{"*" * 50}')
    index = 0
    frames = {}
    for folder_name, paths in files_path.items():
        data = pd.DataFrame()
        for file_path in paths:
            index += 1
            progress = round((index)/sum(num) * 100, 4)
            print(f'{"="*int(progress)}>{index}/{sum(num)}({progress}%) {os.path.basename(file_path)}', end='\r')
            extractor = Extractor(file_path)
            try:
                s = extractor.extract()
                s.name = os.path.basename(file_path)
                data = data.append(s)
            except Exception as e:
                print('file error:', file_path, '\n', e)
        frames[folder_name] = data
    print(end='\n')
    print(f'{"*" * 50}\nfinish parsing, save data to {OUT_PATH}')
    if os.path.isfile(OUT_PATH):
        os.remove(OUT_PATH)
    with pd.ExcelWriter(OUT_PATH) as writer:
        for name, df in frames.items():
            df.to_excel(writer, sheet_name=name)
    print(f'{"*" * 50}\nALL DONE. THANK YOU FOR USING MY PROGRAMME. GOODBYE!\n{"*" * 50}')


# if __name__ == '__main__':
#     main()
