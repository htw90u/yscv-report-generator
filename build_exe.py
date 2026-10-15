"""打包成 Windows 資料夾版（PyInstaller --onedir），並壓成可上傳到 GitHub Release 的 zip，同時印出 SHA-256。

用法：python build_exe.py          （需要 pip install pyinstaller）
輸出：dist/yscv-report-generator/（可直接執行 yscv-report-generator.exe）與 dist/yscv-report-generator-v<版本>-win64.zip
意見反應的設定檔 yscv_feedback_endpoint.py 存在的話會一起打包（網址與 TOKEN 會在 .exe 裡）；不存在就不內建意見反應的寄送。
"""
import hashlib
import os
import shutil
import subprocess
import sys
import zipfile

ROOT = os.path.dirname(os.path.abspath(__file__))
NAME = 'yscv-report-generator'
DATA = ['保育巡查回報單.docx', '塔塔加保育宣導回報單.docx', '生態廁所回報單.docx', 'fonts', 'icons']
sys.path.insert(0, ROOT)


def main():
    import yscv_report_generator as g
    dist = os.path.join(ROOT, 'dist')
    work = os.path.join(ROOT, 'build')
    for path in (os.path.join(dist, NAME), work):
        shutil.rmtree(path, ignore_errors=True)
    info_path = os.path.join(ROOT, 'yscv_build_info.py')
    label = g.version_label()
    with open(info_path, 'w', encoding='utf-8') as f:
        f.write(f'LABEL = {label!r}' + chr(10))          # 打包後的 .exe 沒有 git，把版本標籤（tag.次數-sha）烙進去
    print('烙入版本：', label)
    command = [sys.executable, '-m', 'PyInstaller', '--noconfirm', '--clean', '--onedir', '--windowed',
               '--name', NAME, '--icon', os.path.join(ROOT, 'icons', 'yscv.ico'),
               '--distpath', dist, '--workpath', work, '--specpath', work,
               '--hidden-import', 'yscv_reports', '--hidden-import', 'yscv_updates',
               '--hidden-import', 'yscv_build_info']
    if os.path.exists(os.path.join(ROOT, 'yscv_feedback_endpoint.py')):
        command += ['--hidden-import', 'yscv_feedback_endpoint']
    else:
        print('注意：沒有 yscv_feedback_endpoint.py，打包出來的程式不能寄意見反應')
    for item in DATA:
        command += ['--add-data', f"{os.path.join(ROOT, item)}{os.pathsep}{item if os.path.isdir(os.path.join(ROOT, item)) else '.'}"]
    command.append(os.path.join(ROOT, 'yscv_report_generator.py'))
    try:
        subprocess.run(command, check=True, cwd=ROOT)
    finally:
        if os.path.exists(info_path):
            os.remove(info_path)               # 烙入檔只給打包用，開發環境走 git describe

    folder = os.path.join(dist, NAME)
    for extra in ('README.md', 'LICENSE'):
        shutil.copy(os.path.join(ROOT, extra), folder)
    zip_path = os.path.join(dist, f"{NAME}-v{g.VERSION}-win64.zip")
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for base, _, files in os.walk(folder):
            for name in files:
                full = os.path.join(base, name)
                z.write(full, os.path.join(NAME, os.path.relpath(full, folder)))
    with open(zip_path, 'rb') as f:
        digest = hashlib.sha256(f.read()).hexdigest()
    print(f"\n完成：{zip_path}\nSHA-256：{digest}")


if __name__ == '__main__':
    main()
