# -*- coding: utf-8 -*-
"""用 Windows 自带的 IExpress 把安装包打成单文件 exe（含自检）"""
import io
import os
import shutil
import subprocess
import time

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
PKG = os.path.join(ROOT, '安装包')
WORK = os.path.join(os.environ['TEMP'], 'cet6pkg')
REAL = os.path.join(os.environ['TEMP'], 'feiyu_setup.exe')
FINAL = os.path.join(ROOT, '肥鱼背单词-安装包.exe')
TEST = os.path.join(os.environ['TEMP'], 'cet6_setup_test.exe')
DRY_LOG = os.path.join(os.environ['TEMP'], 'cet6_dry.txt')
IEX = os.path.join(os.environ['SystemRoot'], 'System32', 'iexpress.exe')
TEST = os.path.join(os.environ['TEMP'], 'cet6_setup_test.exe')
DRY_LOG = os.path.join(os.environ['TEMP'], 'cet6_dry.txt')
IEX = os.path.join(os.environ['SystemRoot'], 'System32', 'iexpress.exe')

if os.path.isdir(WORK):
    shutil.rmtree(WORK)
shutil.copytree(PKG, WORK)

files = []
for name in sorted(os.listdir(WORK)):
    p = os.path.join(WORK, name)
    if os.path.isfile(p):
        files.append(name)
    elif os.path.isdir(p):
        for sub in sorted(os.listdir(p)):
            files.append(name + '\\' + sub)


def make_sed(target, app):
    lines = [
        '[Version]', 'Class=IEXPRESS', 'SEDVersion=3', '[Options]',
        'PackagePurpose=InstallApp', 'ShowInstallProgramWindow=1', 'HideExtractAnimation=0',
        'UseLongFileName=1', 'InsideCompressed=0', 'CAB_FixedSize=0', 'CAB_ResvCodeSigning=0',
        'RebootMode=N', 'InstallPrompt=', 'DisplayLicense=', 'FinishMessage=',
        'TargetName=' + target,
        'FriendlyName=肥鱼背单词 安装',
        'AppLaunched=' + app,
        'PostInstallCmd=<None>', 'AdminQuietInstCmd=', 'UserQuietInstCmd=',
        'SourceFiles=SourceFiles', '[Strings]',
    ]
    for i, f in enumerate(files):
        lines.append('FILE%d="%s"' % (i, f))
    lines += ['[SourceFiles]', 'SourceFiles0=' + WORK + '\\', '[SourceFiles0]']
    for i in range(len(files)):
        lines.append('%%FILE%d%%=' % i)
    return '\r\n'.join(lines) + '\r\n'


def build(target, app, tag):
    sed = os.path.join(WORK, tag + '.sed')
    io.open(sed, 'w', encoding='utf-16', newline='\r\n').write(make_sed(target, app))
    if os.path.exists(target):
        os.remove(target)
    r = subprocess.run([IEX, '/N', '/Q', sed], capture_output=True)
    ok = os.path.exists(target) and os.path.getsize(target) > 100000
    print('%-6s exit=%s size=%s ok=%s' % (tag, r.returncode, os.path.getsize(target) if os.path.exists(target) else 0, ok))
    return ok


# 1. 自检包：把 DryRun 的输出重定向到日志
if os.path.exists(DRY_LOG):
    os.remove(DRY_LOG)
app_test = ('powershell.exe -NoProfile -ExecutionPolicy Bypass -Command '
            '"& .\\install-package.ps1 -DryRun *> ' + DRY_LOG + '"')
if build(TEST, app_test, 'test'):
    subprocess.run([TEST], timeout=120)
    if os.path.exists(DRY_LOG):
        txt = io.open(DRY_LOG, encoding='utf-8', errors='ignore').read()
        print('--- 自检输出 ---')
        print(txt.strip()[:600])
    else:
        print('自检没产生日志（可能没跑起来）')

# 2. 正式包：正常安装
app_real = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File install-package.ps1'
build(REAL, app_real, 'real')
print('output:', REAL, os.path.getsize(REAL) if os.path.exists(REAL) else 0)

# 3. 打包成功后复制回项目目录（目标可能被占用，失败重试）
if os.path.exists(FINAL):
    for _ in range(5):
        try:
            os.remove(FINAL)
            break
        except Exception:
            time.sleep(0.6)
try:
    shutil.copyfile(REAL, FINAL)
    print('copied to', FINAL, os.path.getsize(FINAL))
except Exception as e:
    print('copy failed:', e)
    print('installer left at', REAL)
