import os
import glob
import shutil
import argparse

def convert_using_excel_com(src_folder, dst_folder, remove_existing=True):
    """Use Windows Excel COM to open .xls (read-only if needed) and SaveAs .xlsx.

    Requires MS Excel installed and pywin32 (`pip install pywin32`).
    """
    try:
        import win32com.client as win32
    except Exception as e:
        print("❌ 需要安装 pywin32: python -m pip install pywin32")
        raise

    src_folder = os.path.abspath(src_folder)
    dst_folder = os.path.abspath(dst_folder)

    if not os.path.exists(src_folder):
        print(f"❌ 源文件夹不存在: {src_folder}")
        return 0

    if remove_existing and os.path.exists(dst_folder):
        print(f"清空目标文件夹: {dst_folder}")
        shutil.rmtree(dst_folder)
    os.makedirs(dst_folder, exist_ok=True)

    xls_files = glob.glob(os.path.join(src_folder, '**', '*.xls'), recursive=True)
    print(f"发现 {len(xls_files)} 个 .xls 文件，尝试使用 Excel COM 转换（可能需要 Excel 可用）")

    excel = win32.DispatchEx('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False

    converted = 0
    failed = []

    for f in xls_files:
        try:
            rel = os.path.relpath(f, src_folder)
            out_dir = os.path.join(dst_folder, os.path.dirname(rel))
            os.makedirs(out_dir, exist_ok=True)
            out_path = os.path.join(out_dir, os.path.splitext(os.path.basename(f))[0] + '.xlsx')

            print(f"转换: {f} -> {out_path}")

            # Open read-only to avoid lock errors
            wb = excel.Workbooks.Open(f, ReadOnly=True, IgnoreReadOnlyRecommended=True)

            # If workbook is protected and requires password to open, this will raise.
            # Try SaveAs to xlsx (FileFormat=51)
            wb.SaveAs(out_path, FileFormat=51)
            wb.Close(SaveChanges=False)

            converted += 1
        except Exception as e:
            print(f"  ❌ 转换失败: {f} - {e}")
            failed.append((f, str(e)))
            try:
                # attempt to close if partially opened
                wb.Close(SaveChanges=False)
            except Exception:
                pass
            continue

    try:
        excel.Quit()
    except Exception:
        pass

    print(f"\n转换完成：成功 {converted} 个，失败 {len(failed)} 个。目标文件夹: {dst_folder}")
    if failed:
        print("失败列表示例:")
        for f, msg in failed[:10]:
            print(f" - {f}: {msg}")
    return converted, failed


def main():
    parser = argparse.ArgumentParser(description='使用 Excel COM 将 .xls 批量另存为 .xlsx')
    parser.add_argument('src', nargs='?', help='源目录（递归查找 .xls）')
    parser.add_argument('dst', nargs='?', help='目标目录，默认在源目录下创建 converted_com 文件夹')
    parser.add_argument('--no-clear', action='store_true', help='不要清空目标文件夹，追加转换')

    args = parser.parse_args()
    default_src = r"C:\Users\juziq\Desktop\中国县城统计年鉴\2018\县市"
    src = args.src if args.src else default_src
    dst = args.dst if args.dst else os.path.join(src, 'converted_com')

    convert_using_excel_com(src, dst, remove_existing=not args.no_clear)


if __name__ == '__main__':
    main()
