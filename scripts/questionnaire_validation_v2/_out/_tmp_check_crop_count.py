from pathlib import Path
import yaml, openpyxl
cfg=yaml.safe_load(Path(r'scripts/questionnaire_validation_v2/validation_config.yaml').read_text(encoding='utf-8'))
q=Path(cfg['working_dir'])/cfg['questionnaire_file']
wb=openpyxl.load_workbook(q,data_only=True,read_only=True)
ws=wb['Crop list']
hdr=3
headers=[str(c.value or '').strip() for c in ws[hdr]]
ix={h:i for i,h in enumerate(headers)}
sel_col=next((k for k in headers if k.strip().lower() in ('select top 10 crops','select top 10 crops ')),None)
lbl_col=next((k for k in headers if k.strip().lower() in ('label (en)','label (fr)','label (es)','label (ar)')),None)
print('select_col',sel_col)
if sel_col:
    j=ix[sel_col]
    vals=[]
    for r in range(hdr+1, ws.max_row+1):
        v=ws.cell(r,j+1).value
        s=str(v or '').strip()
        if s not in ('','0','0.0','none','nan'):
            vals.append((r,s))
    print('selected_count',len(vals))
    print('sample',vals[:15])
wb.close()
