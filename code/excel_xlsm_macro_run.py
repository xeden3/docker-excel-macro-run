import sys
import shutil
import pythoncom
import win32com.client
import json
# 禁用自动初始化 COM
sys.coinit_flags = 0
def read_excel(xlsm_name, macro_name):
    pythoncom.CoInitialize()  # 初始化 COM
    rtval = {"errcode": 0, "errmsg": ""}
    try:
        # 参数检查
        if not xlsm_name or not macro_name:
            raise ValueError("xlsm_name and macro_name cannot be empty.")

        # 复制原始文件到临时文件
        shutil.copyfile(xlsm_name, r'c:\tmp.xlsm')
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False  # 可选：设置为 True 可以使 Excel 应用程序可见
        # 启用宏自动运行
        excel.Application.AutomationSecurity = 1  # 1 表示低安全级别，允许所有宏运行
        # 打开 Excel 文件
        workbook = excel.Workbooks.Open(r'c:\tmp.xlsm')
        """
        worksheet = workbook.Worksheets(1)  # 获取第一个工作表

        # 获取数据范围
        used_range = worksheet.UsedRange
        rows = used_range.Rows.Count
        columns = used_range.Columns.Count

        # 读取数据
        data = []
        for row in range(1, rows + 1):
            row_data = []
            for col in range(1, columns + 1):
                cell_value = used_range.Cells(row, col).Value
                row_data.append(cell_value)
            data.append(row_data)

        # 输出数据（示例：打印到控制台）
        for row in data:
            print(row)

        """
        excel.Application.Run(macro_name)
        workbook.Save()
        # 关闭 Excel 文件
        workbook.Close()
        excel.Application.Quit()

    except Exception as e:
        rtval["errcode"] = 1
        rtval["errmsg"] = str(e)
    finally:
        pythoncom.CoUninitialize()  # 释放 COM
        # 如果成功则将临时文件复制回原始文件
        if rtval["errcode"] == 0:
            shutil.copyfile(r'c:\tmp.xlsm', xlsm_name)
        return json.dumps(rtval)

if __name__ == "__main__":
    # 使用 param.txt 文件中的内容作为参数
    if len(sys.argv) != 3:
        # print("Usage: python script.py <xlsm_name> <macro_name>")
        rtval = {"errcode": 0, "errmsg": ""}
        rtval["errcode"] = 1
        rtval["errmsg"] = "Usage: python script.py <xlsm_name> <macro_name>"
        sys.exit(json.dumps(rtval))
    xlsm_name = sys.argv[1]
    macro_name = sys.argv[2]
    sys.exit(read_excel(xlsm_name, macro_name))
    
