"""条形码功能自测：生成样例 Excel -> 指定列生成条形码 -> 校验结果。"""
import os
import tempfile

from openpyxl import Workbook

import barcode_tools

WORK = tempfile.mkdtemp(prefix="barcode_test_")
print("工作目录:", WORK)

# 1) 生成样例 Excel（含表头 + 6 行数据）
xlsx = os.path.join(WORK, "样例.xlsx")
wb = Workbook()
ws = wb.active
ws.append(["姓名", "工号", "城市"])
rows = [
    ["张三", "1001", "北京"],
    ["李四", "1002", "上海"],
    ["王五", "1003", "广州"],
    ["赵六", "1004", "深圳"],
    ["钱七", "1005", "杭州"],
    ["孙八", "1006", "成都"],
]
for r in rows:
    ws.append(r)
wb.save(xlsx)
print("样例 Excel 已生成，数据行:", len(rows))

# 2) 读取列名
headers = barcode_tools.get_headers(xlsx)
print("表头:", headers)
assert headers == ["姓名", "工号", "城市"]

# 3) 按"工号"列（索引1）生成 code128 条形码，从第2行开始（首行为表头）
out = os.path.join(WORK, "barcodes")
res = barcode_tools.generate_barcodes(xlsx, 1, out, "code128", True, 2)
print("结果:", res)
assert res["count"] == 6, "应生成 6 张条形码"
assert not res["skipped"], "不应有跳过项"
pngs = [f for f in os.listdir(out) if f.endswith(".png")]
print("图片:", pngs)
assert len(pngs) == 6

print("\n✅ 条形码功能自测通过（列名在第1行场景）！")

# 4) 额外场景：列名在第 2 行（第1行为标题），数据从第 3 行开始
xlsx2 = os.path.join(WORK, "样例_列名在第2行.xlsx")
wb2 = Workbook()
ws2 = wb2.active
ws2.append(["—— 月度报表 ——"])  # 第1行：标题（非表头）
ws2.append(["姓名", "工号", "城市"])  # 第2行：列名
for r in rows:
    ws2.append(r)
wb2.save(xlsx2)

h2 = barcode_tools.get_headers(xlsx2, 2)
print("列名行=2 读取到的表头:", h2)
assert h2 == ["姓名", "工号", "城市"], "应从第2行读取列名"

out2 = os.path.join(WORK, "barcodes2")
res2 = barcode_tools.generate_barcodes(xlsx2, 1, out2, "code128", True, 3)
print("列名在第2行 结果:", res2)
assert res2["count"] == 6 and not res2["skipped"]
assert len([f for f in os.listdir(out2) if f.endswith(".png")]) == 6

print("✅ 列名在第2行场景自测通过！")
print("样例输出目录:", WORK)
