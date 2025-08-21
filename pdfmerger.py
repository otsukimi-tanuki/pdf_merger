import os
import sys
import threading
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

try:
    from PyPDF2 import PdfWriter, PdfReader
except ImportError:
    print("PyPDF2がインストールされていません。")
    print("pip install PyPDF2 でインストールしてください。")
    sys.exit(1)

class PDFMergerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF結合ツール")
        self.root.geometry("700x600")
        
        # 選択されたディレクトリパス
        self.directory_path = tk.StringVar()
        self.output_filename = tk.StringVar(value="merged.pdf")
        
        self.setup_ui()
        
    def setup_ui(self):
        """UIの構築"""
        # メインフレーム
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # ディレクトリ選択セクション
        dir_frame = ttk.LabelFrame(main_frame, text="PDFファイルのディレクトリ選択", padding="10")
        dir_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(dir_frame, text="ディレクトリパス:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(dir_frame, textvariable=self.directory_path, width=50).grid(row=1, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        ttk.Button(dir_frame, text="参照", command=self.browse_directory).grid(row=1, column=1)
        
        dir_frame.columnconfigure(0, weight=1)
        
        # 出力ファイル設定セクション
        output_frame = ttk.LabelFrame(main_frame, text="出力設定", padding="10")
        output_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(output_frame, text="出力ファイル名:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(output_frame, textvariable=self.output_filename, width=30).grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        output_frame.columnconfigure(0, weight=1)
        
        # ファイル一覧表示セクション
        list_frame = ttk.LabelFrame(main_frame, text="見つかったPDFファイル", padding="10")
        list_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # ファイル一覧表示用のTreeview
        self.file_tree = ttk.Treeview(list_frame, columns=("filename", "size"), show="headings", height=8)
        self.file_tree.heading("filename", text="ファイル名")
        self.file_tree.heading("size", text="サイズ")
        self.file_tree.column("filename", width=400)
        self.file_tree.column("size", width=100)
        
        # スクロールバー
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.file_tree.yview)
        self.file_tree.configure(yscrollcommand=scrollbar.set)
        
        self.file_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        # ボタンセクション
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Button(button_frame, text="ファイル一覧更新", command=self.update_file_list).grid(row=0, column=0, padx=(0, 10))
        ttk.Button(button_frame, text="PDF結合実行", command=self.start_merge_thread, style="Accent.TButton").grid(row=0, column=1, padx=(0, 10))
        ttk.Button(button_frame, text="出力フォルダを開く", command=self.open_output_folder).grid(row=0, column=2)
        
        # プログレスバー
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # ログ表示セクション
        log_frame = ttk.LabelFrame(main_frame, text="処理ログ", padding="10")
        log_frame.grid(row=5, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=8, wrap=tk.WORD)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        # グリッドの重み設定
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)
        main_frame.rowconfigure(5, weight=1)
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
    def browse_directory(self):
        """ディレクトリ選択ダイアログを表示"""
        directory = filedialog.askdirectory(title="PDFファイルが格納されているディレクトリを選択")
        if directory:
            self.directory_path.set(directory)
            self.update_file_list()
            
    def update_file_list(self):
        """ファイル一覧を更新"""
        # 既存の項目をクリア
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)
            
        directory = self.directory_path.get()
        if not directory or not os.path.exists(directory):
            return
            
        # PDFファイルを取得してソート
        pdf_files = []
        for file in os.listdir(directory):
            if file.lower().endswith('.pdf'):
                file_path = os.path.join(directory, file)
                file_size = os.path.getsize(file_path)
                pdf_files.append((file, file_size))
        
        # ファイル名でソート
        pdf_files.sort(key=lambda x: x[0])
        
        # TreeViewに追加
        for filename, size in pdf_files:
            size_mb = size / (1024 * 1024)  # MB単位に変換
            self.file_tree.insert("", tk.END, values=(filename, f"{size_mb:.1f} MB"))
            
        self.log(f"見つかったPDFファイル: {len(pdf_files)}個")
        
    def log(self, message):
        """ログメッセージを表示"""
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
    def start_merge_thread(self):
        """PDF結合処理を別スレッドで開始"""
        if not self.directory_path.get():
            messagebox.showerror("エラー", "ディレクトリを選択してください。")
            return
            
        if not self.output_filename.get():
            messagebox.showerror("エラー", "出力ファイル名を入力してください。")
            return
            
        # ボタンを無効化し、プログレスバーを開始
        self.progress.start(10)
        
        # 別スレッドで結合処理を実行
        thread = threading.Thread(target=self.merge_pdfs, daemon=True)
        thread.start()
        
    def merge_pdfs(self):
        """PDFファイルを結合（別スレッドで実行）"""
        try:
            directory = self.directory_path.get()
            output_filename = self.output_filename.get()
            
            if not output_filename.endswith('.pdf'):
                output_filename += '.pdf'
            
            # PDFファイルを取得してソート
            pdf_files = []
            for file in os.listdir(directory):
                if file.lower().endswith('.pdf'):
                    pdf_files.append(file)
            
            pdf_files.sort()
            
            if not pdf_files:
                self.log("エラー: PDFファイルが見つかりません。")
                return
            
            self.log(f"結合開始: {len(pdf_files)}個のファイルを結合します...")
            
            # PDF結合処理
            pdf_writer = PdfWriter()
            total_pages = 0
            
            for i, pdf_file in enumerate(pdf_files, 1):
                file_path = os.path.join(directory, pdf_file)
                self.log(f"処理中 ({i}/{len(pdf_files)}): {pdf_file}")
                
                try:
                    with open(file_path, 'rb') as file:
                        pdf_reader = PdfReader(file)
                        page_count = len(pdf_reader.pages)
                        
                        # 各ページを追加
                        for page_num in range(page_count):
                            page = pdf_reader.pages[page_num]
                            pdf_writer.add_page(page)
                            total_pages += 1
                            
                        self.log(f"  → {page_count}ページを追加")
                        
                except Exception as e:
                    self.log(f"  → 警告: {pdf_file} の読み込みエラー: {e}")
                    continue
            
            # 結合されたPDFを保存
            output_path = os.path.join(directory, output_filename)
            with open(output_path, 'wb') as output_file:
                pdf_writer.write(output_file)
            
            self.log(f"結合完了！")
            self.log(f"保存先: {output_path}")
            self.log(f"総ページ数: {total_pages}ページ")
            
            # 成功ダイアログを表示
            self.root.after(0, lambda: messagebox.showinfo(
                "完了", 
                f"PDF結合が完了しました！\n\n"
                f"ファイル: {output_filename}\n"
                f"総ページ数: {total_pages}ページ"
            ))
            
        except Exception as e:
            self.log(f"エラー: {e}")
            self.root.after(0, lambda: messagebox.showerror("エラー", f"PDF結合中にエラーが発生しました:\n{e}"))
            
        finally:
            # プログレスバーを停止
            self.root.after(0, self.progress.stop)
            
    def open_output_folder(self):
        """出力フォルダを開く"""
        directory = self.directory_path.get()
        if directory and os.path.exists(directory):
            if sys.platform == "win32":
                os.startfile(directory)
            elif sys.platform == "darwin":
                os.system(f"open '{directory}'")
            else:
                os.system(f"xdg-open '{directory}'")
        else:
            messagebox.showwarning("警告", "ディレクトリが選択されていません。")

def main():
    """メイン関数"""
    root = tk.Tk()
    app = PDFMergerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()