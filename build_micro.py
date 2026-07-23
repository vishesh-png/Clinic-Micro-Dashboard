#!/usr/bin/env python3
"""Inject data.json + per-tab rowspec into template.html -> offline/online/repeat.html."""
import os
HERE = os.path.dirname(os.path.abspath(__file__))

def build(tab, rowspec_file, out):
    data = open(os.path.join(HERE, 'data.json')).read()
    rowspec = open(os.path.join(HERE, rowspec_file)).read()
    html = open(os.path.join(HERE, 'template.html')).read()
    html = html.replace('/*__DATA__*/null', data)
    html = html.replace('/*__ROWSPEC__*/null', rowspec)
    html = html.replace("/*__TAB__*/'offline'", "'%s'" % tab)
    title = {'offline': 'Offline', 'online': 'Online', 'repeat': 'Repeat'}[tab]
    html = html.replace('Micro Demand · Offline', 'Micro Demand · ' + title)
    with open(os.path.join(HERE, out), 'w') as f:
        f.write(html)
    print(out, round(os.path.getsize(os.path.join(HERE, out))/1e6, 2), 'MB')

build('offline', 'rowspec_offline.json', 'offline.html')
build('online', 'rowspec_online.json', 'online.html')
build('repeat', 'rowspec_repeat.json', 'repeat.html')
