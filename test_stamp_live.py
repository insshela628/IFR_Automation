"""直接在 AutoCAD 中测试 IFC stamp，不需要 bot。

用法: python test_stamp_live.py
- 自动连接运行中的 AutoCAD（或启动新实例）
- 打开一个 Tatua panel DWG
- 添加 FOR CONSTRUCTION stamp
- 保留 AutoCAD 打开供目视检查
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))

from ifr_automation_v10 import PanelIFCManager
from pathlib import Path
import win32com.client
import time

# 用 Panel Design 里的第一个 DWG 测试
SOURCE = Path(r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)"
              r"\7.New Zealand\Tatua Solar Farm\1. Drawings\1. Native"
              r"\STS1&2 Panel Design& SLD")

# 找一个小的 DWG（非 -01/-06 这种大文件）
test_dwg = None
for f in sorted(SOURCE.glob("TSF-EN-ELE-DRG-10-02.dwg")):
    test_dwg = f
    break
if test_dwg is None:
    # fallback: 任意 DWG
    for f in sorted(SOURCE.glob("*.dwg")):
        if f.stat().st_size < 2_000_000:
            test_dwg = f
            break

if test_dwg is None:
    print("未找到测试 DWG")
    sys.exit(1)

print(f"测试文件: {test_dwg.name} ({test_dwg.stat().st_size / 1024:.0f} KB)")

# 连接 AutoCAD
try:
    acad = win32com.client.GetActiveObject("AutoCAD.Application")
    print("已连接运行中的 AutoCAD")
except Exception:
    print("正在启动 AutoCAD...")
    acad = win32com.client.Dispatch("AutoCAD.Application")
    acad.Visible = True
    for _ in range(60):
        try:
            _ = acad.Documents.Count
            break
        except Exception:
            time.sleep(1)
    time.sleep(3)

# 打开 DWG
print(f"打开 {test_dwg.name}...")
doc = None
for _retry in range(5):
    try:
        doc = acad.Documents.Open(str(test_dwg))
        break
    except Exception as e:
        print(f"  重试 {_retry+1}/5: {e}")
        time.sleep(3)

if doc is None:
    print("无法打开文件")
    sys.exit(1)

# 等待加载
for _ in range(30):
    try:
        _ = doc.ModelSpace.Count
        break
    except Exception:
        time.sleep(1)
time.sleep(2)

# 创建一个临时 PanelIFCManager 来调用 stamp
PROJECT = Path(r"C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)"
               r"\7.New Zealand\Tatua Solar Farm")
mgr = PanelIFCManager(PROJECT, SOURCE, dry_run=False)

# 找 title block
print("查找 title block...")
block_ref, attrs, space = mgr._find_title_block(doc)
if not attrs:
    print("未找到 title block!")
    print("尝试列出所有 block references:")
    ms = doc.ModelSpace
    for i in range(ms.Count):
        ent = ms.Item(i)
        try:
            if ent.EntityName == 'AcDbBlockReference':
                print(f"  '{ent.Name}' at ({ent.InsertionPoint[0]:.1f}, {ent.InsertionPoint[1]:.1f})")
        except:
            pass
    doc.Close(False)
    sys.exit(1)

print(f"Title block 找到: block_ref={block_ref.Name if block_ref else 'None'}, "
      f"attrs={len(attrs)}, space={'有' if space else '无'}")

# 打印 title block bounding box
try:
    mn, mx = block_ref.GetBoundingBox()
    print(f"Title block bbox: ({float(mn[0]):.1f}, {float(mn[1]):.1f}) -> "
          f"({float(mx[0]):.1f}, {float(mx[1]):.1f})")
except Exception as e:
    print(f"无法获取 bbox: {e}")

# 添加 stamp
print("\n添加 FOR CONSTRUCTION stamp...")
mgr._add_ifc_stamp(doc, block_ref, space)

# 不关闭、不保存 — 让用户在 AutoCAD 中检查
print("\n完成! 请在 AutoCAD 中检查 stamp 是否正确显示。")
print("检查完毕后可手动关闭文件（不保存）。")
