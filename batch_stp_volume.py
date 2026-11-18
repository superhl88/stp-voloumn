# -*- coding: utf-8 -*-
import os
import sys
import traceback
import multiprocessing
from datetime import datetime
from concurrent.futures import ProcessPoolExecutor, as_completed

def check_and_import():
    missing = []
    try:
        import OCP
    except ImportError:
        missing.append("cadquery-ocp")
    try:
        import openpyxl
    except ImportError:
        missing.append("openpyxl")
    if missing and not getattr(sys, 'frozen', False):
        print("缺少必要的依赖库，正在尝试自动安装:", ", ".join(missing))
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", *missing])

def calc_step_volume(filepath):
    """
    使用 OpenCascade (OCP) 工业级几何内核读取 STEP 文件并精确计算实体体积。
    返回: (文件名(不含后缀), 完整文件名, 所在路径, 体积_mm3, 体积_cm3, 实体数量, 状态)
    """
    filename = os.path.basename(filepath)
    part_no = os.path.splitext(filename)[0]
    try:
        from OCP.STEPControl import STEPControl_Reader
        from OCP.IFSelect import IFSelect_RetDone
        from OCP.GProp import GProp_GProps
        from OCP.BRepGProp import BRepGProp
        from OCP.TopExp import TopExp_Explorer
        from OCP.TopAbs import TopAbs_SOLID

        reader = STEPControl_Reader()
        status = reader.ReadFile(filepath)
        if status != IFSelect_RetDone:
            return (part_no, filename, filepath, 0.0, 0.0, 0, "读取失败(STEP格式异常)")

        reader.TransferRoots()
        shape = reader.OneShape()
        if shape.IsNull():
            return (part_no, filename, filepath, 0.0, 0.0, 0, "空模型(无几何体)")

        # 统计实体(SOLID)数量
        explorer = TopExp_Explorer(shape, TopAbs_SOLID)
        solid_count = 0
        while explorer.More():
            solid_count += 1
            explorer.Next()

        # 计算精确体积
        props = GProp_GProps()
        BRepGProp.VolumeProperties_s(shape, props)
        vol_mm3 = abs(props.Mass())
        vol_cm3 = vol_mm3 / 1000.0

        if vol_mm3 <= 1e-9:
            return (part_no, filename, filepath, 0.0, 0.0, solid_count, "警告: 体积为0(可能仅含曲面/线框)")

        return (part_no, filename, filepath, round(vol_mm3, 4), round(vol_cm3, 6), solid_count, "成功")

    except Exception as e:
        return (part_no, filename, filepath, 0.0, 0.0, 0, f"异常: {str(e)[:80]}")


def scan_stp_files(root_dir):
    """遍历所在磁盘/目录下的所有 .stp / .step 文件"""
    stp_files = []
    skip_dirs = {
        '$RECYCLE.BIN', 'System Volume Information', 'Windows',
        'Program Files', 'Program Files (x86)', 'ProgramData',
        '.git', '__pycache__', 'build', 'dist'
    }
    
    for dirpath, dirnames, filenames in os.walk(root_dir):
        dirnames[:] = [d for d in dirnames if d not in skip_dirs and not d.startswith('$')]
        for fname in filenames:
            if fname.lower().endswith(('.stp', '.step')):
                stp_files.append(os.path.join(dirpath, fname))
    return stp_files


def save_to_excel(records, output_path):
    """将最多 10000 条记录写入格式化良好的 Excel 文件"""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "STP体积明细"

    headers = ["序号", "料号(文件名)", "完整文件名", "体积 (mm³)", "体积 (cm³)", "实体数量", "状态", "完整路径"]
    ws.append(headers)

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
    ws.row_dimensions[1].height = 26

    data_font = Font(name="微软雅黑", size=10)
    warn_font = Font(name="微软雅黑", size=10, color="C00000")

    for idx, rec in enumerate(records, start=1):
        row_data = [idx, rec[0], rec[1], rec[3], rec[4], rec[5], rec[6], rec[2]]
        ws.append(row_data)
        row_num = idx + 1
        is_ok = (rec[6] == "成功")
        for col_idx in range(1, len(headers) + 1):
            c = ws.cell(row=row_num, column=col_idx)
            c.font = data_font if is_ok else warn_font
            c.border = thin_border
            if col_idx in (1, 6, 7):
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx in (4, 5):
                c.alignment = Alignment(horizontal="right", vertical="center")
                c.number_format = "#,##0.0000" if col_idx == 4 else "#,##0.000000"
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")

    col_widths = [8, 28, 32, 18, 18, 12, 24, 60]
    for i, w in enumerate(col_widths, start=1):
        col_letter = openpyxl.utils.get_column_letter(i)
        ws.column_dimensions[col_letter].width = w

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:H{len(records) + 1}"
    wb.save(output_path)


def main():
    check_and_import()

    if getattr(sys, 'frozen', False):
        current_dir = os.path.dirname(os.path.abspath(sys.executable))
    else:
        current_dir = os.path.dirname(os.path.abspath(__file__))

    drive, _ = os.path.splitdrive(current_dir)
    scan_root = drive + os.sep if drive else current_dir

    print("=" * 60)
    print(f"当前运行位置: {current_dir}")
    print(f"目标扫描磁盘: {scan_root}")
    print("=" * 60)
    print("正在全盘检索 .stp / .step 文件，请稍候...")

    stp_files = scan_stp_files(scan_root)
    total_files = len(stp_files)
    print(f"检索完成！共发现 {total_files} 个 STEP 文件。")

    if total_files == 0:
        print("未找到任何 .stp 或 .step 文件。")
        return

    max_workers = max(1, (os.cpu_count() or 4) - 1)
    print(f"启动 {max_workers} 个并行工作进程开始计算体积...")

    results = [None] * total_files
    completed = 0

    with ProcessPoolExecutor(max_workers=max_workers) as executor:
        future_to_idx = {executor.submit(calc_step_volume, fp): i for i, fp in enumerate(stp_files)}
        for future in as_completed(future_to_idx):
            idx = future_to_idx[future]
            results[idx] = future.result()
            completed += 1
            if completed % 100 == 0 or completed == total_files:
                print(f"进度: {completed} / {total_files} ({completed * 100.0 / total_files:.1f}%)")

    results.sort(key=lambda x: x[0])

    batch_size = 10000
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_folder = os.path.join(current_dir, f"STP体积统计结果_{timestamp}")
    os.makedirs(output_folder, exist_ok=True)

    num_batches = (total_files + batch_size - 1) // batch_size
    for b in range(num_batches):
        start_idx = b * batch_size
        end_idx = min((b + 1) * batch_size, total_files)
        batch_records = results[start_idx:end_idx]
        excel_name = f"STP体积明细_第{b + 1:02d}卷_({start_idx + 1}-{end_idx}).xlsx"
        excel_path = os.path.join(output_folder, excel_name)
        print(f"正在生成: {excel_name} ...")
        save_to_excel(batch_records, excel_path)

    print("=" * 60)
    print(f"全部处理完成！共生成 {num_batches} 个 Excel 文件。")
    print(f"输出目录: {output_folder}")
    print("=" * 60)


if __name__ == "__main__":
    multiprocessing.freeze_support()
    try:
        main()
    except Exception:
        traceback.print_exc()
    input("\n按回车键退出程序...")