from pathlib import Path

ROOT=Path(__file__).resolve().parents[1]
PARTS=ROOT/'source_parts'

def assemble(prefix, destination):
    files=sorted(PARTS.glob(prefix+'.*.part'))
    if not files:
        raise SystemExit(f'No source parts found for {prefix}')
    text=''.join(p.read_text(encoding='utf-8') for p in files)
    dest=ROOT/destination
    dest.parent.mkdir(parents=True,exist_ok=True)
    dest.write_text(text,encoding='utf-8')
    print(f'assembled {destination} from {len(files)} parts')

assemble('template','src/Kaldrix_Data_Analysis_Mastery.template.html')
assemble('builder','build_project.py')
assemble('curriculum','curriculum.json')
