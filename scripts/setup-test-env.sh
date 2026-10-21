#!/usr/bin/env bash
# docx-format-guard 测试环境一键钉死（规格见 docs/test-environment.md）。
# 做四件事：装开源字体包 → 装 fontconfig 替代映射 → 备官方模板 .doc 并转换
# → 对转换产物做样式名归一化。可重复执行（幂等）。
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"

SUDO=""; [ "$(id -u)" -ne 0 ] && SUDO="sudo"

echo "== 1/4 开源字体包（Ubuntu 官方源） =="
# Noto CJK / Liberation 通常已随系统安装；缺了才装。arphic-uming 与 wqy-zenhei
# 不是默认映射目标，装上是为了保留 docs/test-environment.md 里实测过的对照组。
$SUDO apt-get install -y fonts-noto-cjk fonts-liberation \
  fonts-crosextra-carlito fonts-crosextra-caladea \
  fonts-arphic-uming fonts-wqy-zenhei

echo "== 2/4 fontconfig 替代映射 =="
mkdir -p ~/.config/fontconfig/conf.d
cp config/fontconfig/99-docx-format-guard.conf ~/.config/fontconfig/conf.d/
fc-cache -f >/dev/null 2>&1 || true
echo "宋体   -> $(fc-match -f '%{family}' 宋体)"
echo "黑体   -> $(fc-match -f '%{family}' 黑体)"
echo "仿宋   -> $(fc-match -f '%{family}' 仿宋)"

echo "== 3/4 官方模板 .doc -> .docx =="
mkdir -p local-assets/pku
DOC=local-assets/pku/master-2024.doc
if [ ! -f "$DOC" ]; then
  # 北大研究生院“学位工作-下载专区”2024 版硕士学位论文模板（README 有页面链接）
  curl -sL -o "$DOC" "https://grs.pku.edu.cn/docs/2024-02/20240229092055895909.doc"
fi
echo ".doc sha256: $(sha256sum "$DOC" | cut -d' ' -f1)"
soffice --headless --convert-to docx --outdir local-assets/pku "$DOC"

echo "== 4/4 样式名归一化（LO 转换产物 UI 式 -> 内部式，见 docs/test-environment.md） =="
python3 scripts/normalize-template-styles.py local-assets/pku/master-2024.docx

echo
echo "环境就绪。参考版本："
soffice --version | head -1
python3 -c "import sys, importlib.metadata as m; print('python', sys.version.split()[0]); print('python-docx', m.version('python-docx')); print('pymupdf', m.version('pymupdf')); print('pytest', m.version('pytest'))"
echo "跑全套：python -m pytest tests/ -q"
