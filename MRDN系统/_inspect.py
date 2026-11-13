import os
import json
d = json.load(open(r'dist\MRDN系统\data\data.json', encoding='utf-8'))
for f in d.get('flows', []):
    print('FLOW', f.get('name'), 'steps=', len(f.get('steps', [])))
    for s in f.get('steps', []):
        inp = s.get('input') or {}
        tool = s.get('tool')
        print('  ', s.get('id'), s.get('name'), '| input keys=', list(inp.keys()))
        print('       refs=', inp.get('refs'), '| file=', bool(inp.get('file')), '| fileRequired=', inp.get('fileRequired'))
        print('       tool=', json.dumps(tool, ensure_ascii=False))
