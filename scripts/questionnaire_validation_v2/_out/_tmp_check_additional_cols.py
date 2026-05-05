from pathlib import Path
import yaml, openpyxl
cfg=yaml.safe_load(Path(r'scripts/questionnaire_validation_v2/validation_config.yaml').read_text(encoding='utf-8'))
q=Path(cfg['working_dir'])/cfg['questionnaire_file']
wb=openpyxl.load_workbook(q,data_only=True,read_only=True)
ws=next((wb[n] for n in wb.sheetnames if n.strip().lower()=='additional information'),None)
headers=[str(c.value or '').strip() for c in ws[2]]
ix={h:i for i,h in enumerate(headers)}
print('headers',headers)
for col in ['Original','Replacement (EN)','Replacement (FR)','Replacement']:
    if col not in ix:
        continue
    j=ix[col]
    n=0
    for r in range(3, ws.max_row+1):
        if str(ws.cell(r,j+1).value or '').strip():
            n+=1
    print(col,'nonblank',n)
wb.close()
