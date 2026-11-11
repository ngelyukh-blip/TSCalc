"""Summarize an audit snapshot, preserving the distinction between attributes and geometry."""
import argparse
import collections
import json
from decimal import Decimal
from pathlib import Path

def number(value):
    return Decimal(str(value).strip().replace(',', '.'))

def summarize(snapshot):
    entities = snapshot['entities']
    panels = [e for e in entities if e.get('effective_name') == 'щит']
    joists = [e for e in entities if e.get('effective_name') == 'лага']
    def attrs(e):
        return {a['tag']: a['value'] for a in e.get('attributes', [])}
    lengths = [number(attrs(e)['ДЛИНА']) for e in joists]
    areas = [number(attrs(e)['ПЛОЩАДЬ']) for e in panels]
    return {
        'drawing': snapshot['dwg_name'],
        'dbmod_at_read': snapshot['dbmod'],
        'assumed_units': 'mm — explicitly confirmed by user',
        'entity_types':dict(collections.Counter(e['type'] for e in entities)),
        'panels_insert_count':len(panels),
        'joists_insert_count':len(joists),
        'joist_length_m_from_attributes':str(sum(lengths, Decimal(0))/1000),
        'panel_area_m2_from_attributes':str(sum(areas, Decimal(0))),
        'unvalidated_geometry':True,
        'objects_with_xdata':sum(e.get('xdata_raw','nil') != 'nil' for e in entities),
        'curves_by_layer': dict(collections.Counter(e['layer'] for e in entities if e['type'] in ('CIRCLE','LWPOLYLINE'))),
        'notes':[
            'Attribute totals have not been independently reconciled with clipped geometry.',
            'Insert counts are not purchase quantities; source drawing and clipping remain unchanged.',
            'Snapshot string formatting limits precision; obtain numeric geometry before final specification.'
        ]
    }

if __name__ == '__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('snapshot',type=Path)
    parser.add_argument('--output',type=Path)
    args=parser.parse_args()
    result=summarize(json.loads(args.snapshot.read_text(encoding='utf-8')))
    text=json.dumps(result,ensure_ascii=False,indent=2)
    if args.output:
        args.output.write_text(text,encoding='utf-8')
    print(text)
