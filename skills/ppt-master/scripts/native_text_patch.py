#!/usr/bin/env python3
"""Replace one complete DrawingML text run, preserving every other part byte.

Deliberately limited to one unique full a:t run in one explicitly selected shape.
Split runs, tables, math and complex revisions belong to the upstream roundtrip.
"""
import argparse
import hashlib
import json
from pathlib import Path
import re
import sys
import zipfile
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

NS = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}


def patch(source, output, slide, shape_id, old, new):
    import posixpath
    source, output = Path(source).resolve(), Path(output).resolve()
    if source == output or output.exists():
        raise ValueError('Choose a new output file; source replacement is forbidden')
    if not old or not new or old == new or new.strip() != new:
        raise ValueError('Require distinct nonempty old/new text without new boundary whitespace')
    ET.fromstring(('<t>' + escape(new) + '</t>').encode('utf-8'))
    source_hash = hashlib.sha256(source.read_bytes()).hexdigest()
    with zipfile.ZipFile(source) as zin:
        names = zin.namelist()
        if len(names) != len(set(names)) or zin.testzip():
            raise ValueError('Source ZIP integrity failure')
        if any(n.startswith('_xmlsignatures/') for n in names):
            raise ValueError('Signed presentations require a signature-aware workflow')
        presentation = ET.fromstring(zin.read('ppt/presentation.xml'))
        roster = presentation.findall('./p:sldIdLst/p:sldId', NS)
        if not 1 <= slide <= len(roster):
            raise ValueError('Slide number outside presentation roster')
        rid = roster[slide-1].get('{'+NS['r']+'}id')
        rels = ET.fromstring(zin.read('ppt/_rels/presentation.xml.rels'))
        matches = [r for r in rels if r.get('Id') == rid and r.get('TargetMode') != 'External']
        if len(matches) != 1:
            raise ValueError('Ambiguous or missing slide relationship')
        target = matches[0].get('Target', '')
        part = posixpath.normpath(target.lstrip('/') if target.startswith('/') else 'ppt/'+target)
        if not part.startswith('ppt/slides/') or '..' in part.split('/'):
            raise ValueError('Unexpected slide part')
        original = zin.read(part)
        tree = ET.fromstring(original)
        targets = [s for s in tree.findall('.//p:sp', NS)
                   if (s.find('./p:nvSpPr/p:cNvPr', NS) is not None and
                       s.find('./p:nvSpPr/p:cNvPr', NS).get('id') == str(shape_id))]
        if len(targets) != 1:
            raise ValueError('Require one explicit plain native shape')
        texts = targets[0].findall('./p:txBody/a:p/a:r/a:t', NS)
        if len([t for t in texts if t.text == old]) != 1:
            raise ValueError('Old text must match one complete ordinary run in the target shape')
        # No XML serialization: preserve namespaces, attributes, runs, autofit and geometry.
        raw = original.decode('utf-8')
        pattern = re.compile(r'(<(?P<prefix>[A-Za-z_][\w.-]*:)?t(?:\s[^>]*)?>)(?P<body>[^<]*)(</(?P=prefix)t>)')
        hits = []
        for hit in pattern.finditer(raw):
            try:
                decoded = ET.fromstring('<t>'+hit.group('body')+'</t>').text
            except ET.ParseError:
                continue
            if decoded == old:
                hits.append(hit)
        if len(hits) != 1:
            raise ValueError('Text is lexically ambiguous; use the upstream roundtrip route')
        h = hits[0]
        changed = (raw[:h.start('body')] + escape(new) + raw[h.end('body'):]).encode('utf-8')
        after = ET.fromstring(changed)
        before_run = next(t for t in tree.findall('.//a:t', NS) if t.text == old)
        before_run.text = new
        if ET.tostring(tree) != ET.tostring(after):
            raise ValueError('Unexpected XML change outside the selected text run')
        output.parent.mkdir(parents=True, exist_ok=True)
        try:
            with zipfile.ZipFile(output, 'w') as zout:
                for info in zin.infolist():
                    zout.writestr(info, changed if info.filename == part else zin.read(info.filename))
        except BaseException:
            output.unlink(missing_ok=True)
            raise
    if hashlib.sha256(source.read_bytes()).hexdigest() != source_hash:
        raise RuntimeError('Source unexpectedly changed')
    return {'status': 'pass', 'slide': slide, 'shape_id': shape_id,
            'changed_part': part, 'old': old, 'new': new,
            'source_sha256': source_hash,
            'output_sha256': hashlib.sha256(output.read_bytes()).hexdigest(),
            'limits': ['One unique complete ordinary a:t run only; no reflow or font fitting is performed.',
                       'Inspect the output rendering; longer text can overflow even when XML is exact.']}


def main():
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument('source'); ap.add_argument('output')
    ap.add_argument('--slide',type=int,required=True); ap.add_argument('--shape-id',type=int,required=True)
    ap.add_argument('--old',required=True); ap.add_argument('--new',required=True)
    ap.add_argument('--report')
    a=ap.parse_args()
    try:
        report=patch(a.source,a.output,a.slide,a.shape_id,a.old,a.new)
    except (ValueError, KeyError, ET.ParseError, zipfile.BadZipFile) as ex:
        ap.exit(2, str(ex)+'\n')
    if a.report:
        rp=Path(a.report); rp.parent.mkdir(parents=True,exist_ok=True)
        rp.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(report,ensure_ascii=False,indent=2))


if __name__ == '__main__':
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    main()
