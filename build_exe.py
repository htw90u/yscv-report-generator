"""打包成 Windows 資料夾版（PyInstaller --onedir），並壓成可上傳到 GitHub Release 的 zip，同時印出 SHA-256。

用法：python build_exe.py          （需要 pip install pyinstaller）
輸出：build/stage/玉山保育志工回報單產生器/（可直接執行「玉山保育志工回報單產生器.exe」）與 dist/yscv-report-generator-v<版本>-win64.zip
（zip 檔名維持英文，GitHub 會把附件檔名裡的中文拿掉）
意見反應的設定檔 yscv_feedback_endpoint.py 存在的話會一起打包（網址與 TOKEN 會在 .exe 裡）；不存在就不內建意見反應的寄送。
"""
import hashlib
import os
import shutil
import subprocess
import sys
import zipfile

ROOT = os.path.dirname(os.path.abspath(__file__))
NAME = '玉山保育志工回報單產生器'                 # 資料夾與 .exe 的名稱（中文）
ZIP_PREFIX = 'yscv-report-generator'     # zip 檔名用英文，上傳 GitHub Release 時中文檔名會被改掉
DATA = ['保育巡查回報單.docx', '塔塔加保育宣導回報單.docx', '生態廁所回報單.docx', 'fonts', 'icons']
sys.path.insert(0, ROOT)


def main():
    import yscv_report_generator as g
    dist = os.path.join(ROOT, 'dist')                 # 只放最後的 zip
    work = os.path.join(ROOT, 'build')                # 暫存與打包出來的資料夾（stage），每次重新建立
    stage = os.path.join(work, 'stage')               # 資料夾版放這裡：就算 dist 裡舊版的程式還開著也不會卡住
    shutil.rmtree(work, ignore_errors=True)
    os.makedirs(dist, exist_ok=True)
    info_path = os.path.join(ROOT, 'yscv_build_info.py')
    label = g.version_label()
    with open(info_path, 'w', encoding='utf-8') as f:
        f.write(f'LABEL = {label!r}' + chr(10))          # 打包後的 .exe 沒有 git，把版本標籤（tag.次數-sha）烙進去
    print('烙入版本：', label)
    command = [sys.executable, '-m', 'PyInstaller', '--noconfirm', '--clean', '--onedir', '--windowed',
               '--name', NAME, '--icon', os.path.join(ROOT, 'icons', 'yscv.ico'),
               '--distpath', stage, '--workpath', work, '--specpath', work,
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

    folder = os.path.join(stage, NAME)
    for extra in ('README.md', 'LICENSE'):
        shutil.copy(os.path.join(ROOT, extra), folder)
    manual = os.path.join(ROOT, 'docs', '使用說明.pdf')            # 給使用者的 PDF 使用說明（python tools/make_manual.py 產生）
    if os.path.exists(manual):
        shutil.copy(manual, folder)
    else:
        print('注意：沒有 docs/使用說明.pdf，zip 裡不會有使用說明')
    zip_path = os.path.join(dist, f"{ZIP_PREFIX}-v{g.VERSION}-win64.zip")
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
