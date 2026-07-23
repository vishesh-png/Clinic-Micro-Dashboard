#!/usr/bin/env python3
"""Inject data.json + rowspec + tab into template.html -> <tab>.html.

Usage:
    python3 build_micro.py            # builds offline.html (only tab implemented so far)
"""
import os, json
HERE = os.path.dirname(os.path.abspath(__file__))

def build(tab, rowspec_file, out):
    data = open(os.path.join(HERE, 'data.json')).read()
    rowspec = open(os.path.join(HERE, rowspec_file)).read()
    html = open(os.path.join(HERE, 'template.html')).read()
    html = html.replace('/*__DATA__*/null', data)
    html = html.replace('/*__ROWSPEC__*/null', rowspec)
    html = html.replace("/*__TAB__*/'offline'", "'%s'" % tab)
    with open(os.path.join(HERE, out), 'w') as f:
        f.write(html)
    print(out, round(os.path.getsize(os.path.join(HERE, out))/1e6, 2), 'MB')

build('offline', 'rowspec_offline.json', 'offline.html')
