# -*- coding: utf-8 -*-
"""Оценка объёма работы в страницах по норме 1800 знаков.

    python3 tools/objem.py referat/grammatika.md
"""
import re, sys, pathlib

t = pathlib.Path(sys.argv[1]).read_text(encoding='utf-8')
c = re.sub(r'<!--\w+-->', '', t)
c = re.sub(r'[#*`|>\-]', '', c)
c = re.sub(r'\n{2,}', '\n', c)
print(f"{len(c)} знаков = {len(c)/1800:.1f} стр. (норма 1800 знаков на страницу)")
