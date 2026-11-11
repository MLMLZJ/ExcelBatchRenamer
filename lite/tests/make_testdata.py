
import os, shutil
from openpyxl import Workbook

base = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "testdata")
if os.path.isdir(base):
    shutil.rmtree(base)
src = os.path.join(base, "源文件夹")
os.makedirs(src)
out = os.path.join(base, "输出文件夹")
os.makedirs(out)

for n in ["FILE10001.pdf", "FILE10002.pdf", "FILE10003.pdf", "FILE10004.pdf", "FILE10005.pdf", "报告 001.pdf"]:
    open(os.path.join(src, n), "wb").write(b"%PDF-1.4\n% test fixture\n")

wb = Workbook()
ws = wb.active
ws.title = "清单"
ws.append(["编号", "文件编号"])
for r in [
    (1, "FILE10001"),
    (2, "FILE10002"),
    (3, "FILE10003"),
    (4, "FILE10004"),
    (5, "FILE10005.pdf"),
    (6, "FILE10006"),          # 缺失
    (7, "报告 001"),            # 带空格
    (1, "FILE10001"),          # 与第 1 行新名字重复
    ("A08", "FILE10001"),
]:
    ws.append(list(r))
wb.save(os.path.join(base, "清单.xlsx"))

batch = os.path.join(base, "批量测试")
os.makedirs(batch)
for n in ["final report.pdf", "final report 2.pdf", "IMG_001.JPG", "readme.txt"]:
    open(os.path.join(batch, n), "wb").write(b"x")

print("fixture ok")
