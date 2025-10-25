# main.py (修改 login_feishu 方法)
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import pandas as pd
from feishu_auth import start_authorize_flow, exchange_code_for_token
from feishu_sheet import append_to_sheet, col_num_to_letter

class FeishuApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Feishu 自动化工具")
        self.root.geometry("500x600")
        self.access_token = None

        tk.Label(root, text="飞书自动化上传工具", font=("Arial", 16, "bold")).pack(pady=15)

        tk.Button(root, text="1️⃣ 登录飞书授权", command=self.login_feishu, width=30, height=2).pack(pady=10)
        tk.Button(root, text="2️⃣ 选择本地Excel并上传", command=self.select_excel, width=30, height=2).pack(pady=10)

        self.status_label = tk.Label(root, text="状态：未授权", fg="red", font=("Arial", 12))
        self.status_label.pack(pady=20)

    def login_feishu(self):
        # 启动授权（非阻塞）
        start_authorize_flow()
        messagebox.showinfo("请授权", "浏览器已打开，请完成登录授权。")
        self.check_authorization()

    def check_authorization(self):
        """每1秒检测是否已获取到 code"""
        from localserver import last_code
        if last_code:
            try:
                token = exchange_code_for_token(last_code)
                self.access_token = token
                self.status_label.config(text="状态：已授权 ✅", fg="green")
                messagebox.showinfo("成功", "飞书授权成功！")
                print(f"access_token: {token}")
            except Exception as e:
                messagebox.showerror("错误", f"获取token失败：{e}")
        else:
            # 1 秒后再次检查
            self.root.after(1000, self.check_authorization)

    def select_excel(self):
        if not self.access_token:
            messagebox.showwarning("警告", "请先完成飞书授权！")
            return

        file_path = filedialog.askopenfilename(title="选择Excel文件", filetypes=[("Excel Files", "*.xlsx *.xls")])
        if not file_path:
            return

        df = pd.read_excel(file_path)
        messagebox.showinfo("Excel列信息", f"读取到 {len(df.columns)} 列：{', '.join(df.columns)}")

        spreadsheet_token = simpledialog.askstring("输入表格Token", "请输入飞书云文档 spreadsheet_token：")
        sheet_name = simpledialog.askstring("输入Sheet名称", "请输入Sheet名称（例如 Sheet1）")
        start_cell = simpledialog.askstring("输入起始单元格", "请输入起始单元格（例如 A2）：")

        try:
            # 自动计算结束单元格（跳过表头）
            num_rows = len(df) - 1
            num_cols = len(df.columns)

            # 自动解析列字母与行号
            start_col_letter = ''.join(filter(str.isalpha, start_cell)).upper()
            start_row_str = ''.join(filter(str.isdigit, start_cell))
            if not start_col_letter or not start_row_str:
                raise ValueError("起始单元格格式错误，应形如 A2")

            start_row = int(start_row_str)
            end_col_letter = col_num_to_letter(
                num_cols + (ord(start_col_letter) - ord('A'))
            )
            end_row = start_row + num_rows - 1
            range_ref = f"{sheet_name}!{start_col_letter}{start_row}:{end_col_letter}{end_row}"

            # ✅ 去掉表头，只上传数据部分
            values = df.iloc[1:].values.tolist()

            print(f"📘 上传范围: {range_ref}")
            print(f"📘 上传数据预览（前2行）: {values[:2]}")

            success = append_to_sheet(self.access_token, spreadsheet_token, range_ref, values)
            if success:
                messagebox.showinfo("完成", f"成功追加到飞书云文档！\n范围：{range_ref}")
            else:
                messagebox.showerror("失败", "上传失败，请检查控制台输出。")

        except Exception as e:
            messagebox.showerror("错误", f"处理表格时出错：\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = FeishuApp(root)
    root.mainloop()
