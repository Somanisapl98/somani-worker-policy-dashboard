from pathlib import Path
p=Path(__file__).resolve().parent/'input'
files=sorted([x for x in p.glob('OT DT*.xlsx') if not x.name.startswith('~$')],key=lambda x:x.stat().st_mtime,reverse=True)
for x in files[1:]:
    x.unlink(); print('Deleted old worker input:',x.name)
