#!/usr/bin/env python3
"""Read-only finished-PPTX audit, strictly scoped comparison, vector PDF cropping.

Original enhancement for this candidate; no Anthropic code is incorporated.
The audit reports evidence and review targets, never an aesthetic score.
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import json
import math
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET

from pptx_opc_validation import canonical_opc_part_path, resolve_internal_opc_target

NS = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
}
SHAPES = {f"{{{NS['p']}}}{name}" for name in ('sp', 'pic', 'graphicFrame', 'cxnSp', 'grpSp')}


def sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def c14n(node: ET.Element) -> str:
    # Prefix rewriting makes namespace-prefix-only rewrites irrelevant here.
    return ET.canonicalize(ET.tostring(node, encoding='unicode'), rewrite_prefixes=True)


def shape_id(node: ET.Element) -> str | None:
    # The immediate non-visual properties belong to this shape; never borrow
    # an ID from a child of a group.
    for child in node:
        if child.tag.rsplit('}', 1)[-1].startswith('nv'):
            props = child.find('p:cNvPr', NS)
            if props is not None:
                return props.get('id')
    return None


def read_package(path: Path) -> tuple[dict[str, bytes], list[dict]]:
    issues = []
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        dupes = [name for name, count in Counter(names).items() if count > 1]
        if dupes:
            issues.append({'code': 'duplicate_members', 'parts': dupes})
        bad = z.testzip()
        if bad:
            issues.append({'code': 'crc_failure', 'part': bad})
        parts = {name: z.read(name) for name in names if not name.endswith('/')}
    keys = {}
    for name in parts:
        key = canonical_opc_part_path(name)
        if key is None:
            issues.append({'code': 'invalid_part_path', 'part': name})
        elif key in keys:
            issues.append({'code': 'canonical_part_collision', 'parts': [keys[key], name]})
        else:
            keys[key] = name
    return parts, issues


def parse_xml(parts: dict[str, bytes], issues: list[dict]) -> dict[str, ET.Element]:
    trees = {}
    for name, content in parts.items():
        if name.endswith(('.xml', '.rels')):
            try:
                trees[name] = ET.fromstring(content)
            except ET.ParseError as exc:
                issues.append({'code': 'invalid_xml', 'part': name, 'message': str(exc)})
    return trees


def roster(trees: dict[str, ET.Element]) -> list[str]:
    root = trees.get('ppt/presentation.xml')
    rels = trees.get('ppt/_rels/presentation.xml.rels')
    if root is None or rels is None:
        return []
    targets = {r.get('Id'): resolve_internal_opc_target('ppt/_rels/presentation.xml.rels', r.get('Target', '')) for r in rels}
    result = []
    for entry in root.findall('p:sldIdLst/p:sldId', NS):
        part = targets.get(entry.get(f"{{{NS['r']}}}id"))
        if part:
            result.append(part)
    return result


def shape_map(root: ET.Element) -> tuple[dict[str, ET.Element], list[str]]:
    shapes, duplicates = {}, []
    for node in root.iter():
        if node.tag in SHAPES:
            sid = shape_id(node)
            if sid is None:
                duplicates.append('(missing)')
            elif sid in shapes:
                duplicates.append(sid)
            else:
                shapes[sid] = node
    return shapes, duplicates


def visible_bounds(node: ET.Element, scale=(1.0, 1.0), offset=(0.0, 0.0)):
    """Yield axis-aligned shape boxes, including nested group transforms.

    Rounded bounding boxes account for shape rotation, but not native effects,
    text ink, clipping, or inherited placeholder positions.
    """
    if node.tag not in SHAPES:
        for child in node:
            yield from visible_bounds(child, scale, offset)
        return
    xfrm = node.find('p:spPr/a:xfrm', NS)
    if xfrm is None:
        xfrm = node.find('p:xfrm', NS)
    if node.tag == f"{{{NS['p']}}}grpSp":
        xfrm = node.find('p:grpSpPr/a:xfrm', NS)
        if xfrm is not None:
            off, ext = xfrm.find('a:off', NS), xfrm.find('a:ext', NS)
            coff, cext = xfrm.find('a:chOff', NS), xfrm.find('a:chExt', NS)
            if all(x is not None for x in (off, ext, coff, cext)):
                sx = int(ext.get('cx', '0')) / max(1, int(cext.get('cx', '0')))
                sy = int(ext.get('cy', '0')) / max(1, int(cext.get('cy', '0')))
                child_scale = (scale[0] * sx, scale[1] * sy)
                child_offset = (
                    offset[0] + scale[0] * (int(off.get('x', '0')) - sx * int(coff.get('x', '0'))),
                    offset[1] + scale[1] * (int(off.get('y', '0')) - sy * int(coff.get('y', '0'))),
                )
                for child in node:
                    yield from visible_bounds(child, child_scale, child_offset)
                return
        for child in node:
            yield from visible_bounds(child, scale, offset)
        return
    if xfrm is not None:
        off, ext = xfrm.find('a:off', NS), xfrm.find('a:ext', NS)
        if off is not None and ext is not None:
            x = offset[0] + int(off.get('x', '0')) * scale[0]
            y = offset[1] + int(off.get('y', '0')) * scale[1]
            w = int(ext.get('cx', '0')) * scale[0]
            h = int(ext.get('cy', '0')) * scale[1]
            angle = math.radians(int(xfrm.get('rot', '0')) / 60000)
            rw, rh = abs(w * math.cos(angle)) + abs(h * math.sin(angle)), abs(w * math.sin(angle)) + abs(h * math.cos(angle))
            yield shape_id(node), (x + w / 2 - rw / 2, y + h / 2 - rh / 2, rw, rh)


def audit(path: Path) -> dict:
    parts, errors = read_package(path)
    trees = parse_xml(parts, errors)
    keys = {canonical_opc_part_path(name) for name in parts}
    external = []
    for name, root in trees.items():
        if not name.endswith('.rels'):
            continue
        if root.tag != f"{{{NS['rel']}}}Relationships":
            errors.append({'code': 'invalid_relationship_namespace', 'part': name})
            continue
        seen = set()
        for rel in root:
            rid, target, typ = rel.get('Id'), rel.get('Target'), rel.get('Type')
            mode = rel.get('TargetMode', '').lower()
            if not rid or rid in seen or not target or not typ or mode not in ('', 'internal', 'external'):
                errors.append({'code': 'invalid_relationship', 'part': name, 'attributes': rel.attrib})
            seen.add(rid)
            if mode == 'external':
                external.append({'part': name, 'id': rid, 'target': target, 'type': typ})
                continue
            dest = resolve_internal_opc_target(name, target or '')
            if dest not in keys:
                errors.append({'code': 'dangling_relationship', 'part': name, 'id': rid, 'target': target})
    # Resolve relationship IDs used by XML, including images and hyperlinks.
    # Relationship existence alone is insufficient when the XML names a stale ID.
    for name, root in trees.items():
        if not name.endswith('.xml'):
            continue
        directory, _, filename = name.rpartition('/')
        rel_name = f'{directory}/_rels/{filename}.rels' if directory else f'_rels/{filename}.rels'
        rel_tree = trees.get(rel_name)
        rel_ids = {r.get('Id') for r in rel_tree} if rel_tree is not None else set()
        for node in root.iter():
            for key, value in node.attrib.items():
                if key.startswith(f"{{{NS['r']}}}") and value not in rel_ids:
                    errors.append({'code': 'unresolved_relationship_id', 'part': name, 'id': value})
    for required in ('[Content_Types].xml', '_rels/.rels', 'ppt/presentation.xml'):
        if required not in parts:
            errors.append({'code': 'missing_required_part', 'part': required})
    p = trees.get('ppt/presentation.xml')
    size = p.find('p:sldSz', NS) if p is not None else None
    width, height = (int(size.get('cx', '0')), int(size.get('cy', '0'))) if size is not None else (0, 0)
    if not width or not height:
        errors.append({'code': 'invalid_canvas'})
    pages = roster(trees)
    if not pages:
        errors.append({'code': 'empty_or_invalid_slide_roster'})
    slide_records, concerns = [], []
    for page, part in enumerate(pages, 1):
        root = trees.get(part)
        if root is None:
            errors.append({'code': 'missing_roster_slide', 'slide': page, 'part': part})
            continue
        smap, duplicate_ids = shape_map(root)
        if duplicate_ids:
            errors.append({'code': 'duplicate_or_missing_shape_id', 'slide': page, 'ids': duplicate_ids})
        kinds = Counter(node.tag.rsplit('}', 1)[-1] for node in smap.values())
        texts = [''.join(node.itertext()) for node in root.findall('.//a:t', NS)]
        boxes = []
        for sid, (x, y, w, h) in visible_bounds(root):
            if x < -9144 or y < -9144 or x + w > width + 9144 or y + h > height + 9144:
                item = {'slide': page, 'shape_id': sid, 'code': 'bounds_suspect', 'bbox_emu': [round(v) for v in (x, y, w, h)]}
                concerns.append(item)
                boxes.append(item)
        slide_records.append({
            'slide': page, 'part': part, 'shape_count_including_groups': len(smap),
            'native_shapes': kinds['sp'], 'connectors': kinds['cxnSp'], 'groups': kinds['grpSp'],
            'pictures': kinds['pic'], 'graphic_frames': kinds['graphicFrame'],
            'text_shapes': sum(bool(n.findall('.//a:t', NS)) for n in smap.values() if n.tag != f"{{{NS['p']}}}grpSp"),
            'charts': len(root.findall('.//c:chart', NS)), 'tables': len(root.findall('.//a:tbl', NS)),
            'native_math_objects': len(root.findall('.//m:oMath', NS)),
            'text': '\n'.join(texts), 'bounds_concerns': boxes,
        })
    return {
        'schema': 'ppt-master.enhanced-audit.v1', 'file': str(path.resolve()),
        'sha256': sha(path.read_bytes()), 'bytes': path.stat().st_size,
        'status': 'failed' if errors else ('review-needed' if concerns else 'pass'),
        'canvas_emu': [width, height], 'package_parts': len(parts), 'slide_count': len(slide_records),
        'errors': errors, 'bounds_concerns': concerns, 'external_relationships': external, 'slides': slide_records,
        'limits': [
            'Counts establish object presence, not practical editability of every mark or semantic correctness.',
            'Geometry is a review hint; effects, grouped rotation, text ink, and inherited placeholders need rendered inspection.',
            'No check for text wrapping, overlap, aesthetics, or application repair dialogs is inferred from XML.',
        ],
    }


def masked(root: ET.Element, ids: set[str]) -> ET.Element:
    root = copy.deepcopy(root)
    for parent in root.iter():
        for index, child in enumerate(list(parent)):
            if child.tag in SHAPES and shape_id(child) in ids:
                replacement = ET.Element('QA-ALLOWED-SHAPE', {'id': shape_id(child) or ''})
                replacement.tail = child.tail
                parent.remove(child)
                parent.insert(index, replacement)
    return root


def compare(before: Path, after: Path, allowed: list[str]) -> dict:
    bp, errors = read_package(before)
    ap, aerrors = read_package(after)
    errors.extend(aerrors)
    bt, at = parse_xml(bp, errors), parse_xml(ap, errors)
    br, ar = roster(bt), roster(at)
    if not br or not ar:
        errors.append({'code': 'empty_or_invalid_slide_roster'})
    allowed_map: dict[str, set[str]] = {}
    for request in allowed:
        match = re.fullmatch(r'([1-9][0-9]*):([0-9]+)', request)
        if not match or int(match[1]) > len(br):
            errors.append({'code': 'invalid_allow_shape', 'value': request})
            continue
        allowed_map.setdefault(br[int(match[1]) - 1], set()).add(match[2])
    changed_parts = [name for name in sorted(bp.keys() | ap.keys()) if bp.get(name) != ap.get(name)]
    added, removed = sorted(ap.keys() - bp.keys()), sorted(bp.keys() - ap.keys())
    if br != ar:
        errors.append({'code': 'slide_roster_changed', 'before': br, 'after': ar})
    for name in changed_parts:
        if name not in allowed_map:
            errors.append({'code': 'unapproved_part_changed', 'part': name})
    shape_changes = []
    approved_actual = []
    for page, part in enumerate(br, 1):
        if part not in bt or part not in at:
            continue
        bmap, bd = shape_map(bt[part])
        amap, ad = shape_map(at[part])
        if bd or ad:
            errors.append({'code': 'duplicate_or_missing_shape_id', 'slide': page, 'before': bd, 'after': ad})
        permitted = allowed_map.get(part, set())
        for sid in permitted:
            if sid not in bmap or sid not in amap:
                errors.append({'code': 'allowed_shape_missing', 'slide': page, 'id': sid})
        for sid in sorted(bmap.keys() | amap.keys(), key=lambda s: int(s) if s.isdigit() else -1):
            bnode, anode = bmap.get(sid), amap.get(sid)
            different = bnode is None or anode is None or c14n(bnode) != c14n(anode)
            if different:
                approved = sid in permitted
                shape_changes.append({'slide': page, 'id': sid, 'allowed_directly': approved,
                                      'before_sha256_c14n': sha(c14n(bnode).encode()) if bnode is not None else None,
                                      'after_sha256_c14n': sha(c14n(anode).encode()) if anode is not None else None})
                if approved:
                    approved_actual.append(f'{page}:{sid}')
        if c14n(masked(bt[part], permitted)) != c14n(masked(at[part], permitted)):
            errors.append({'code': 'unapproved_slide_structure_changed', 'slide': page, 'part': part})
    return {
        'schema': 'ppt-master.scoped-compare.v1', 'status': 'pass' if not errors else 'failed',
        'before': str(before.resolve()), 'after': str(after.resolve()),
        'allowed_shapes': allowed, 'allowed_shapes_actually_changed': approved_actual,
        'changed_parts': changed_parts, 'added_parts': added, 'removed_parts': removed,
        'changed_shapes': shape_changes,
        'unchanged_part_payloads': len(bp.keys() & ap.keys()) - sum(n in bp and n in ap for n in changed_parts),
        'errors': errors,
        'limits': [
            'Unchanged package member payloads are compared byte for byte; ZIP compression and container headers may differ.',
            'Within an allowed slide, removing only allowed shape subtrees must leave canonical XML identical.',
            'An allowed group authorizes its descendants. Choose the smallest shape scope the requested edit needs.',
            'Shape XML preservation does not prove identical rendering; inspect both exports.',
        ],
    }


def crop_pdf(source: Path, output: Path, page: int, bbox: list[float]) -> dict:
    from pypdf import PdfReader, PdfWriter
    from pypdf.generic import RectangleObject
    if source.resolve() == output.resolve() or output.exists():
        raise ValueError('Output must be a new file and cannot replace the source.')
    reader = PdfReader(source)
    if not 1 <= page <= len(reader.pages):
        raise ValueError('Page is 1-based and must exist.')
    x0, y0, x1, y1 = bbox
    if not all(math.isfinite(v) for v in bbox) or x1 <= x0 or y1 <= y0:
        raise ValueError('bbox requires finite x0 y0 x1 y1 with positive area.')
    selected = reader.pages[page - 1]
    mb = list(map(float, selected.mediabox))
    if x0 < mb[0] or y0 < mb[1] or x1 > mb[2] or y1 > mb[3]:
        raise ValueError('bbox must lie within the source MediaBox.')
    if selected.rotation:
        raise ValueError('Rotated PDF pages need a separately verified coordinate transform; normalize the source first.')
    original_contents = selected.get_contents()
    before_content = original_contents.get_data() if original_contents is not None else b''
    selected.mediabox = RectangleObject(bbox)
    selected.cropbox = RectangleObject(bbox)
    selected.trimbox = RectangleObject(bbox)
    selected.bleedbox = RectangleObject(bbox)
    selected.artbox = RectangleObject(bbox)
    writer = PdfWriter()
    writer.add_page(selected)
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open('wb') as stream:
        writer.write(stream)
    check = PdfReader(output).pages[0].get_contents()
    after_content = check.get_data() if check is not None else b''
    if before_content != after_content:
        raise RuntimeError('Unexpected change to PDF content stream.')
    return {
        'schema': 'ppt-master.vector-crop.v1', 'status': 'pass', 'source': str(source.resolve()),
        'output': str(output.resolve()), 'source_page': page, 'bbox_pdf_points': bbox,
        'content_stream_preserved': True, 'content_stream_sha256': sha(before_content),
        'limits': ['Changes PDF page boxes without rasterization; no automatic content-bound detection.',
                   'Cropping is not redaction. Content outside the displayed page can remain recoverable.',
                   'Source vectors stay vectors; pre-existing raster images remain raster.'],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest='command', required=True)
    p = sub.add_parser('audit')
    p.add_argument('pptx', type=Path)
    p.add_argument('--strict-bounds', action='store_true', help='Return failure also for geometric bounds concerns.')
    p.add_argument('--report', type=Path)
    p = sub.add_parser('compare')
    p.add_argument('before', type=Path)
    p.add_argument('after', type=Path)
    p.add_argument('--allow-shape', action='append', default=[], metavar='SLIDE:ID')
    p.add_argument('--report', type=Path)
    p = sub.add_parser('crop-pdf')
    p.add_argument('source', type=Path)
    p.add_argument('output', type=Path)
    p.add_argument('--page', type=int, required=True)
    p.add_argument('--bbox', type=float, nargs=4, required=True, metavar=('X0', 'Y0', 'X1', 'Y1'))
    p.add_argument('--report', type=Path)
    args = parser.parse_args()
    try:
        if args.command == 'audit':
            report = audit(args.pptx)
            result = 1 if report['errors'] or (args.strict_bounds and report['bounds_concerns']) else 0
        elif args.command == 'compare':
            report = compare(args.before, args.after, args.allow_shape)
            result = 0 if report['status'] == 'pass' else 1
        else:
            report = crop_pdf(args.source, args.output, args.page, args.bbox)
            result = 0
    except Exception as exc:
        report = {'status': 'failed', 'error': f'{type(exc).__name__}: {exc}'}
        result = 1
    rendered = json.dumps(report, ensure_ascii=False, indent=2)
    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(rendered + '\n', encoding='utf-8')
    print(rendered)
    return result


if __name__ == '__main__':
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    raise SystemExit(main())
