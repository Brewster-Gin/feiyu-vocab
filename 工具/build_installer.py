# -*- coding: utf-8 -*-
"""用 IExpress 打正式包：在 ASCII 临时目录构建（中文/空格路径会让它静默失败）"""
import io
import os
import shutil
import subprocess
import time

PKG = r'D:\MyProject\Big Fat Fish Vocabulary\CET6-Plan\安装包'
SRC = r'C:\Users\Lenovo\AppData\Local\Temp\cet6pkg2'
OUT = r'C:\Users\Lenovo\AppData\Local\Temp\feiyu_setup.exe'
FINAL = r'D:\MyProject\Big Fat Fish Vocabulary\CET6-Plan\肥鱼背单词-安装包.exe'

if os.path.isdir(SRC):
    shutil.rmtree(SRC)
shutil.copytree(PKG, SRC)
# 不要把预览图、截图、日志和构建残留打进包
for junk in os.listdir(SRC):
    low = junk.lower()
    if low == 'build.sed' or low.endswith('.png') or low.endswith('.log'):
        os.remove(os.path.join(SRC, junk))
files = [f for f in sorted(os.listdir(SRC)) if os.path.isfile(os.path.join(SRC, f))]
print('files:', len(files))
for f in files:
    print('  -', f)

if 'setup-gui.ps1' not in files or 'setup-launch.vbs' not in files:
    raise SystemExit('缺少 setup-gui.ps1 / setup-launch.vbs')

lines = [
    '[Version]', 'Class=IEXPRESS', 'SEDVersion=3', '[Options]',
    'PackagePurpose=InstallApp', 'ShowInstallProgramWindow=1', 'HideExtractAnimation=0',
    'UseLongFileName=1', 'InsideCompressed=0', 'CAB_FixedSize=0', 'CAB_ResvCodeSigning=0',
    'RebootMode=N', 'InstallPrompt=', 'DisplayLicense=', 'FinishMessage=',
    'TargetName=' + OUT, 'FriendlyName=FeiyuVocab Setup',
    'AppLaunched=wscript.exe setup-launch.vbs',
    'PostInstallCmd=<None>', 'AdminQuietInstCmd=', 'UserQuietInstCmd=',
    'SourceFiles=SourceFiles', '[Strings]',
]
for i, f in enumerate(files):
    lines.append('FILE%d="%s"' % (i, f))
lines += ['[SourceFiles]', 'SourceFiles0=' + SRC + '\\', '[SourceFiles0]']
for i in range(len(files)):
    lines.append('%%FILE%d%%=' % i)

sed = os.path.join(SRC, 'build.sed')
io.open(sed, 'w', encoding='utf-16', newline='\r\n').write('\r\n'.join(lines) + '\r\n')

if os.path.exists(OUT):
    os.remove(OUT)
r = subprocess.run([os.path.join(os.environ['SystemRoot'], 'System32', 'iexpress.exe'), '/N', '/Q', sed],
                   capture_output=True, text=True, errors='ignore')
size = os.path.getsize(OUT) if os.path.exists(OUT) else 0
print('iexpress exit', r.returncode, 'size', size)
if size > 100000:
    # 校验：CAB 头里能搜到关键文件名
    blob = io.open(OUT, 'rb').read()
    for probe in (b'setup-gui.ps1', b'setup-launch.vbs', b'install-package.ps1', b'build_plan.py'):
        print('  contains', probe.decode(), probe in blob)
    if os.path.exists(FINAL):
        for _ in range(5):
            try:
                os.remove(FINAL)
                break
            except Exception:
                time.sleep(0.6)
    shutil.copyfile(OUT, FINAL)
    print('final:', FINAL, os.path.getsize(FINAL))
else:
    print('build failed')
