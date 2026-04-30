import json
from pathlib import Path
for p in [r"scripts/questionnaire_validation_v2/questionnaire_validation_geopoll.ipynb", r"scripts/questionnaire_validation_v2/questionnaire_validation_kobo.ipynb"]:
    nb=json.loads(Path(p).read_text(encoding='utf-8'))
    print('\n===',p,'===')
    for i,c in enumerate(nb['cells']):
        if c.get('cell_type')!='code':
            continue
        s=''.join(c.get('source',[]))
        if '_write_config_snapshot' in s:
            idx=s.index('_write_config_snapshot')
            print('cell',i)
            print(s[max(0,idx-350):idx+1200].encode('ascii','backslashreplace').decode('ascii'))
            break
