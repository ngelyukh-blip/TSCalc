"""Build the private role-instruction package; never include source drawings or prices."""
from pathlib import Path
import json
import shutil
import zipfile

ROOT = Path(__file__).resolve().parents[1]
PACKAGE = ROOT / 'plugins' / 'tscalc-team'

def build():
    refs = PACKAGE / 'references'
    refs.mkdir(exist_ok=True)
    for name in ('DECISIONS.md', 'TEAM.md', 'BACKLOG.md', 'SOURCE_AUDIT.md', 'NORMATIVE_DOCUMENTS.md'):
        shutil.copyfile(ROOT / 'docs' / name, refs / name)
    manifest = json.loads((PACKAGE / 'plugin.json').read_text(encoding='utf-8'))
    assert manifest['name'] == PACKAGE.name
    assert manifest['$schema'] == 'https://agent-plugins.org/schemas/1.0.0/plugin.schema.json'
    interface = manifest['extensions']['com.openai']['interface']
    assert len(interface['shortDescription']) <= 30
    assert interface['category'] == 'Developer Tools'
    assert isinstance(interface['capabilities'], list)
    skills = list((PACKAGE / 'skills').glob('*/SKILL.md'))
    assert len(skills) == 3
    names = set()
    for skill in skills:
        # JSON is a YAML 1.2 subset. The frontmatter deliberately uses that
        # subset so it can be parsed with the standard library, without YAML deps.
        content = skill.read_text(encoding='utf-8')
        parts = content.split('---', 2)
        assert not parts[0].strip() and len(parts) == 3
        meta = json.loads(parts[1])
        assert meta['name'] == skill.parent.name and meta['description'].strip()
        assert meta['name'] not in names
        names.add(meta['name'])
        import re
        for target in re.findall(r'\]\(([^)]+)\)', parts[2]):
            resolved = (skill.parent / target).resolve()
            assert resolved.is_relative_to(PACKAGE.resolve()) and resolved.is_file(), target
    outdir = ROOT / 'dist'
    outdir.mkdir(exist_ok=True)
    archive = outdir / f"tscalc-team-{manifest['version']}.zip"
    files = sorted(p for p in PACKAGE.rglob('*') if p.is_file())
    with zipfile.ZipFile(archive, 'w', zipfile.ZIP_DEFLATED) as z:
        for f in files:
            assert f.suffix.lower() not in {'.dwg', '.xlsx', '.jsonl', '.dll'}
            z.write(f, f.relative_to(PACKAGE).as_posix())
    with zipfile.ZipFile(archive) as z:
        assert z.testzip() is None
        assert 'plugin.json' in z.namelist()
        assert len([n for n in z.namelist() if n.endswith('/SKILL.md')]) == 3
        json.loads(z.read('plugin.json'))
    print(json.dumps({'archive':str(archive), 'files':len(files), 'roles':sorted(names)}, ensure_ascii=False))

if __name__ == '__main__':
    build()
