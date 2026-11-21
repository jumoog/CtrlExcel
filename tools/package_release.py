"""Packs a CtrlExcelReader release: one zip per platform, laid out like a
WinCC OA project, so it unpacks straight into the project directory
(README and licenses go to doc/CtrlExcelReader/, so a project's own README
is not overwritten), plus SHA256SUMS.txt. See AGENTS.md, "Releasing".

    python tools/package_release.py OUT_DIR
        [--windows build/RelWithDebInfo/CtrlExcelReader.dll]
        [--debian12 path/to/stripped/CtrlExcelReader.so]
        [--rhel10 path/to/stripped/CtrlExcelReader.so]

The version comes from CMakeLists.txt (project(... VERSION)). Pass the
exact library files that were tested; the Linux ones stripped of debug info.
The other files come from the working tree, so HEAD must be the release tag
v<version> and they must be unchanged (--allow-untagged skips that check,
for trying the script out). The zip entries carry the commit's date.
"""

import argparse
import hashlib
import os
import re
import subprocess
import zipfile

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

COMMON = [
    ('scripts/libs/classes/ExcelWorkbook.ctl', 'scripts/libs/classes/ExcelWorkbook.ctl'),
    ('README.md', 'doc/CtrlExcelReader/README.md'),
    ('LICENSE', 'doc/CtrlExcelReader/LICENSE'),
    ('THIRD-PARTY-LICENSES.txt', 'doc/CtrlExcelReader/THIRD-PARTY-LICENSES.txt'),
]

# option -> (zip name part, name of the library in the zip)
PLATFORMS = {
    'windows': ('Windows-x64', 'bin/CtrlExcelReader.dll'),
    'debian12': ('Linux-Debian12-x64', 'bin/CtrlExcelReader.so'),
    'rhel10': ('Linux-RHEL10-x64', 'bin/CtrlExcelReader.so'),
}


def version():
    with open(os.path.join(REPO, 'CMakeLists.txt'), encoding='utf-8') as f:
        match = re.search(r'project\(CtrlExcelReader VERSION ([0-9.]+)\)', f.read())
    if not match:
        raise SystemExit('CMakeLists.txt: project(CtrlExcelReader VERSION x.y.z) not found')
    return match.group(1)


def git(*args):
    return subprocess.run(['git', *args], cwd=REPO, check=True, capture_output=True, text=True).stdout.strip()


def check_release_tree(ver):
    """HEAD is tagged v<ver> and the files packed from the tree are unchanged."""
    tags = git('tag', '--points-at', 'HEAD').split()
    if f'v{ver}' not in tags:
        raise SystemExit(f'HEAD is not the release tag v{ver} (tags at HEAD: {tags or "none"})')
    changed = git('status', '--porcelain', '--', *(source for source, _ in COMMON), 'CMakeLists.txt')
    if changed:
        raise SystemExit('files packed from the tree differ from the tag:\n' + changed)


def commit_date():
    # The tagged commit's date: reproducible, unlike the time of packing.
    date = git('log', '-1', '--format=%cd', '--date=format:%Y %m %d %H %M %S')
    return tuple(int(part) for part in date.split())


def add(z, source, name, mode, date):
    # Unix modes, so unzip on Linux keeps the library executable.
    info = zipfile.ZipInfo(name, date_time=date)
    info.compress_type = zipfile.ZIP_DEFLATED
    info.create_system = 3
    info.external_attr = (0o100000 | mode) << 16
    with open(source, 'rb') as f:
        z.writestr(info, f.read())


def main():
    parser = argparse.ArgumentParser(description=__doc__.split('\n\n')[0])
    parser.add_argument('out_dir')
    parser.add_argument('--allow-untagged', action='store_true',
                        help='pack although HEAD is not the release tag or the files are changed')
    for option in PLATFORMS:
        parser.add_argument('--' + option, metavar='LIBRARY')
    args = parser.parse_args()

    ver = version()
    if not args.allow_untagged:
        check_release_tree(ver)
    date = commit_date()
    os.makedirs(args.out_dir, exist_ok=True)
    sums = []
    for option, (platform, library_name) in PLATFORMS.items():
        library = getattr(args, option)
        if not library:
            continue
        zip_name = f'CtrlExcelReader-{ver}-{platform}.zip'
        path = os.path.join(args.out_dir, zip_name)
        with zipfile.ZipFile(path, 'w') as z:
            add(z, library, library_name, 0o755, date)
            for source, name in COMMON:
                add(z, os.path.join(REPO, source), name, 0o644, date)
        with open(path, 'rb') as f:
            sums.append(f'{hashlib.sha256(f.read()).hexdigest()}  {zip_name}')
        print(f'{zip_name}: {os.path.getsize(path)} bytes')
    if not sums:
        raise SystemExit('no library given (--windows, --debian12, --rhel10)')
    with open(os.path.join(args.out_dir, 'SHA256SUMS.txt'), 'w', encoding='ascii', newline='\n') as f:
        f.write('\n'.join(sums) + '\n')


if __name__ == '__main__':
    main()
