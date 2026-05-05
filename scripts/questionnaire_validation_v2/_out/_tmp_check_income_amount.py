from pathlib import Path
import yaml, openpyxl

cfg = yaml.safe_load(Path(r'scripts/questionnaire_validation_v2/validation_config.yaml').read_text(encoding='utf-8'))
wd = Path(cfg['working_dir'])
cur = wd / cfg['questionnaire_file']
tdir = Path(cfg['templates_dir'])
lang = cfg['language'].upper()

cands = [p for p in tdir.glob('*.xlsx') if 'geopoll' in p.name.lower() and f'_{lang}_' in p.name.upper()]
cands = sorted(cands, key=lambda p: p.stat().st_mtime, reverse=True)
ref = cands[0] if cands else None

print('CURRENT', cur)
print('REFERENCE', ref)

Q='income_third_amount'
COLS=['Skip Pattern','Default skip patterns & conditional','Default skip patterns & conditional ','Specify skip pattern variable (from blue text)']

def get_row(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb['survey']
    hdr = 3
    headers=[str(c.value or '').strip() for c in ws[hdr]]
    ix={h:i for i,h in enumerate(headers)}
    out={'row':None}
    for r in range(hdr+1, ws.max_row+1):
        qn=str(ws.cell(r, ix['Q Name']+1).value or '').strip()
        if qn==Q:
            out['row']=r
            for c in COLS:
                j=ix.get(c)
                if j is not None:
                    out[c]=str(ws.cell(r,j+1).value or '').strip()
            break
    wb.close()
    return out

for p in [cur, ref]:
    o=get_row(p)
    print('\nFILE', p.name, 'row', o['row'])
    for c in COLS:
        if c in o:
            print(' ',c,':', o[c])
