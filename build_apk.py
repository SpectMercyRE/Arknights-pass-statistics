#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""一键构建安卓 APK

做三件事：
  1. 把 arknights-pass/ 的前端源码（index.html / data.js / images）同步到
     app/www/ 与 app/android/app/src/main/assets/public/（Capacitor 打包用的是这两处）
  2. 调用 Gradle 构建 debug APK
  3. 把产物复制到项目根目录，文件名带上 index.html 里的版本号

用法：python build_apk.py
"""
import os
import re
import shutil
import subprocess
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(ROOT, 'arknights-pass')
APP = os.path.join(ROOT, 'app')
DEST_WWW = os.path.join(APP, 'www')
DEST_ASSETS = os.path.join(APP, 'android', 'app', 'src', 'main', 'assets', 'public')


def read_version():
    """从 index.html 的关于页读当前版本号（回填文案时一起更新）"""
    html = open(os.path.join(SRC, 'index.html'), encoding='utf-8').read()
    m = re.search(r'about-ver">v([\d.]+)', html)
    return m.group(1) if m else '0.0'


def sync(src, dst):
    """把前端源码同步到目标目录（index.html / data.js / images）"""
    os.makedirs(dst, exist_ok=True)
    for name in ('index.html', 'data.js'):
        shutil.copy2(os.path.join(src, name), os.path.join(dst, name))
    img_src = os.path.join(src, 'images')
    img_dst = os.path.join(dst, 'images')
    if os.path.exists(img_src):
        # 逐个文件覆盖（不做整目录删除重建，避免误删 + 批量删除保护）
        n = 0
        for base, _dirs, files in os.walk(img_src):
            rel = os.path.relpath(base, img_src)
            target = img_dst if rel == '.' else os.path.join(img_dst, rel)
            os.makedirs(target, exist_ok=True)
            for f in files:
                shutil.copy2(os.path.join(base, f), os.path.join(target, f))
                n += 1
        print('  图片 %d 个' % n)
    print('  同步 ->', os.path.relpath(dst, ROOT))


def main():
    ver = read_version()
    print('构建版本: v%s' % ver)

    print('[1/3] 同步前端源码')
    sync(SRC, DEST_WWW)
    sync(SRC, DEST_ASSETS)

    print('[2/3] Gradle 构建（首次较慢，请耐心）')
    gradlew = os.path.join(APP, 'android', 'gradlew.bat')
    if os.name != 'nt':
        gradlew = os.path.join(APP, 'android', 'gradlew')
    r = subprocess.run([gradlew, 'assembleDebug', '--no-daemon'],
                       cwd=os.path.join(APP, 'android'))
    if r.returncode != 0:
        print('构建失败 ✗ 退出码', r.returncode)
        return r.returncode

    print('[3/3] 复制产物')
    apk = os.path.join(APP, 'android', 'app', 'build', 'outputs', 'apk', 'debug', 'app-debug.apk')
    out = os.path.join(ROOT, '通行认证收集册-v%s.apk' % ver)
    shutil.copy2(apk, out)
    print('完成 ✓', out)
    return 0


if __name__ == '__main__':
    sys.exit(main())
