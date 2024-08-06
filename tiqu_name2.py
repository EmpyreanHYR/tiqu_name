import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
import traceback


def extract_names(file_path, start_cell, col_letter):
    try:
        # 计算要跳过的行数和列数
        start_row = int(start_cell[1:]) - 1
        use_col = col_letter.upper()

        # 读取Excel文件
        df = pd.read_excel(file_path, header=None, skiprows=start_row, usecols=use_col)
        print(f"读取的DataFrame: \n{df}")

        # 提取“学生姓名”列
        names = df.iloc[:, 0].dropna().tolist()
        print(f"提取的姓名: {names}")

        # 在名字后面加上中文逗号，每三个名字换一行
        output_lines = []
        line = ''
        for i, name in enumerate(names):
            if (i + 1) % 3 == 0:
                line += name + '，'
                output_lines.append(line)
                line = ''
            else:
                line += name + '，'

        # 如果最后一行有未满3个名字，仍然加入到输出中
        if line:
            output_lines.append(line)

        # 输出到txt文件，文件名与Excel文件相同，路径相同
        output_file_path = os.path.splitext(file_path)[0] + '_output.txt'
        with open(output_file_path, 'w', encoding='utf-8') as f:
            for line in output_lines:
                f.write(line + '\n')

        return output_file_path

    except Exception as e:
        messagebox.showerror("错误", f"发生错误: {str(e)}")
        print(traceback.format_exc())
        return None


def main():
    try:
        # 创建Tkinter主窗口
        root = tk.Tk()
        root.withdraw()

        messagebox.showinfo("提示", "请按下Enter回车键以开始使用本程序提取以指定单元格开始的姓名")

        # 弹出文件选择对话框
        file_path = filedialog.askopenfilename(title="选择Excel文件", filetypes=[("Excel文件", "*.xlsx *.xls")])
        if not file_path:
            messagebox.showinfo("提示", "未选择任何文件。")
            return

        print(f"已选择文件: {file_path}")

        # 获取用户输入的起始单元格和列号
        start_cell = simpledialog.askstring("输入", "请输入起始单元格(例如B3):")
        col_letter = simpledialog.askstring("输入", "请输入列字母(例如B):")

        if not start_cell or not col_letter:
            messagebox.showinfo("提示", "未输入起始单元格或列字母。")
            return

        output_file_path = extract_names(file_path, start_cell, col_letter)

        if output_file_path:
            # 创建并显示一个包含提取完成信息的窗口
            info_window = tk.Toplevel(root)
            info_window.title("完成")
            tk.Label(info_window, text=f"提取姓名已完成\n导出的文件已保存到:\n{output_file_path}").pack(padx=20,
                                                                                                        pady=20)
            tk.Button(info_window, text="关闭", command=root.quit).pack(pady=10)
            info_window.mainloop()

    except Exception as e:
        messagebox.showerror("错误", f"发生错误: {str(e)}")
        print(traceback.format_exc())

    finally:
        root.quit()


if __name__ == '__main__':
    main()
