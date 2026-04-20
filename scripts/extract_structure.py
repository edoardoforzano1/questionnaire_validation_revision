import json
import sys

def get_notebook_structure(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        nb = json.load(f)
    
    print(f"\n{'='*90}")
    print(f"FILE: {filepath.split('/')[-1]}")
    print(f"{'='*90}")
    
    cells = nb.get('cells', [])
    print(f"\nTotal cells: {len(cells)}")
    
    cell_summary = {"code": 0, "markdown": 0}
    
    for i, cell in enumerate(cells):
        cell_type = cell.get('cell_type', 'unknown')
        source_lines = cell.get('source', [])
        
        if isinstance(source_lines, list):
            source = ''.join(source_lines)
        else:
            source = source_lines
        
        source_stripped = source.strip()
        
        if cell_type == 'markdown':
            cell_summary['markdown'] += 1
            first_line = source_stripped.split('\n')[0][:70]
            print(f"\n[Cell {i+1:2d}] {cell_type.upper():10s} | {first_line}")
        elif cell_type == 'code':
            cell_summary['code'] += 1
            # Get first meaningful line
            lines = [l for l in source.split('\n') if l.strip()]
            if lines:
                first = lines[0][:70]
            else:
                first = "(empty)"
            print(f"\n[Cell {i+1:2d}] {cell_type.upper():10s} | {first}")
    
    print(f"\n\nCell Summary: {cell_summary['code']} code cells, {cell_summary['markdown']} markdown cells")

files = [
    'GeoPoll_questionnaire_validation.ipynb',
    'newQ_test.ipynb',
    'newQ_testv2.ipynb'
]

for f in files:
    try:
        get_notebook_structure(f)
    except Exception as e:
        print(f"Error processing {f}: {e}")
