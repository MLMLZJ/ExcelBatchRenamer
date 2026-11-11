
import os, shutil
from openpyxl import Workbook

base = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "testdata")
edge = os.path.join(base, "边界测试")
if os.path.isdir(edge):
    shutil.rmtree(edge)
src = os.path.join(edge, "文件")
os.makedirs(src)
for n in ["A 001.pdf", "B 002.pdf", "C 003.pdf", "D 004.pdf"]:
    open(os.path.join(src, n), "wb").write(b"x")

wb = Workbook()
ws = wb.active
ws.title = "数据"
ws.append(["编号", "文件编号"])          # 表头
ws.append([1.0, "A 001"])                 # 数字 1
ws.append([2, "B 002.pdf"])               # 文本带扩展名
ws.append([None, None])                   # 空行
ws.append([3, "C 003"])                   # 空行后继续
ws.append([4, "D 004"])                   # 同名校验
ws["A7"] = "注释行"                        # 单列行
ws2 = wb.create_sheet("说明")
ws2.append(["第二个工作表，不应被读取"])
wb.save(os.path.join(edge, "边界.xlsx"))
print("edge fixture ok", edge)