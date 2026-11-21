"""Remove OOXML personal metadata only. Never patch VBA binary or algorithms."""
from pathlib import Path
import sys, zipfile, re

def sanitize(path):
    path = Path(path)
    with zipfile.ZipFile(path) as z:
        contents = [(i, z.read(i.filename)) for i in z.infolist()]
    output = path.with_suffix('.sanitized.xlam')
    if output.exists():
        raise RuntimeError('Sanitization output already exists')
    with zipfile.ZipFile(output, 'w') as z:
        for info, data in contents:
            if info.filename in ['docProps/core.xml', 'docProps/app.xml', 'xl/workbook.xml']:
                if info.filename.startswith('docProps/'):
                    for name in [b'creator',b'lastModifiedBy',b'Company',b'Manager']:
                        pattern = rb'(<(?:[\w]+:)?'+name+rb'\b[^>]*>).*?(</(?:[\w]+:)?'+name+rb'>)'
                        data = re.sub(pattern, rb'\1\2', data, flags=re.S)
                else:
                    # Excel's build-directory hint is unnecessary for this add-in.
                    # Preserve namespace declarations and mc:Ignorable prefixes.
                    data = re.sub(rb'<(?:[\w]+:)?absPath\b[^>]*/>', b'', data)
            z.writestr(info, data)
    output.replace(path)

if __name__ == '__main__':
    sanitize(sys.argv[1])
    print('Core OOXML personal metadata removed; VBA binary untouched.')
