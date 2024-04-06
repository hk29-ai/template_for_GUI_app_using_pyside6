import os
import sys
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.cm as cm
import japanize_matplotlib
plt.rcParams['font.size'] = 15 # グラフの基本フォントサイズの設定
from PySide6.QtWidgets import QApplication, QFileDialog, QMainWindow, QPushButton, QVBoxLayout, QWidget, QLabel, QTextEdit, QListWidget, QAbstractItemView


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("アプリの雛形") 
        self.x_name = 'x'
        self.y_name = 'y'

        # ファイル選択ボタンを作成
        self.select_file_button = QPushButton("Excelファイルを選択")
        self.select_file_button.clicked.connect(self.open_file_dialog)

        # ファイルパスを表示するラベル
        self.file_path_label = QLabel() 
     
        # シート名を表示するリストウィジット
        self.sheet_list = QListWidget()
        self.sheet_list.setSelectionMode(QAbstractItemView.MultiSelection)

        # シート選択ボタン
        self.select_sheet_button = QPushButton("データ処理の実行")
        self.select_sheet_button.clicked.connect(self.add_selected_sheets_to_list)

        # レイアウト 
        layout = QVBoxLayout()
        layout.addWidget(self.select_file_button)
        layout.addWidget(self.file_path_label) 
        layout.addWidget(self.sheet_list)
        layout.addWidget(self.select_sheet_button)

        # ウィジェットを作成し、メインウィンドウに設定 
        widget = QWidget()
        widget.setLayout(layout)
        self.setCentralWidget(widget)

    # ファイルの選択
    def open_file_dialog(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, "excelファイルを選択", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
        if file_path:
            # シートリストを初期化
            self.sheet_list.clear()
            self.file_path_label.setText(file_path)

            # シート名の一覧を取得する
            excel_file = pd.ExcelFile(file_path)
            sheet_names = excel_file.sheet_names
            self.sheet_list.addItems(sheet_names)
            
            # ファイルパスをインスタンス変数化
            self.selected_excel_file = file_path
            
            # フォルダパスを取得してインスタンス変数化
            self.folder_path = os.path.dirname(file_path)
            
            # ファイル名の取得
            file_name = os.path.basename(file_path)
            # ファイルパスから、拡張子なしのファイル名と拡張子の取得
            file_name_wo_extension, extension = os.path.splitext(file_name)
            self.file_name = file_name_wo_extension

    # 選択したシートに対して実行
    def add_selected_sheets_to_list(self):
        # 選択したシートをリストへ格納する
        selected_items = self.sheet_list.selectedItems()
        self.selected_sheet_names = [item.text() for item in selected_items]
        print('選択されたーシート', self.selected_sheet_names) 

        # シート毎に処理する（グラフを作成して、画像ファイルに保存する）
        data_dict = {} # 各シートデータをひとつにするための辞書
        for i, select_sheet_name in enumerate(self.selected_sheet_names, start=1):
            _df = pd.read_excel(self.selected_excel_file,
                                sheet_name = select_sheet_name,
                                usecols = "A:B", # AからB列までを選択
                                engine = "openpyxl"
            )        
            # データを抽出
            x_data = _df[self.x_name].values
            y_data = _df[self.y_name].values
            
            # データを辞書へ格納
            data_dict[select_sheet_name] = (x_data, y_data)
            
            # 散布図
            fig = plt.figure(figsize=(6,4))
            plt.scatter(x_data, y_data)
            # 折れ線図
            plt.plot(x_data, y_data)
            
            if select_sheet_name == 'ガンマ関数':
                plt.yscale('log')
            plt.grid()
            plt.xlabel(f"{self.x_name}")
            plt.ylabel(f"{self.y_name}")
            plt.title(f"{select_sheet_name}")
            
            # 保存するファイルパス
            _save_jpg_path = f"{self.folder_path}/{str(i).zfill(2)}_{select_sheet_name}.jpg"
            
            # 画像ファイルに出力
            plt.savefig(_save_jpg_path, bbox_inches="tight")
            print(f"save {_save_jpg_path}")
            plt.clf()

        ### シート別のデータを1つに結合する
        # data_dictからデータを取得してリストに追加
        rows = []
        for select_sheet_name, (x_data, y_data) in data_dict.items():
            for x, y in zip(x_data, y_data):
                rows.append([select_sheet_name, x, y])
        
        # データフレーム作成
        df = pd.DataFrame(rows, columns=["select_sheet_name", "x_data", "y_data"])
        save_csv_path = f"{self.folder_path}/{self.file_name}.csv"
        df.to_csv(save_csv_path, index=False, encoding="shift-jis")
        print(f"save {save_csv_path}")
        
        
        ### ガンマ分布の確率密度関数のデータだけ抽出してグラフ化する
        x_name = 'x_data'
        y_name = 'y_data'
        label = 'select_sheet_name'
        legend_name = ''
        title = 'ガンマ分布の確率密度関数'

        # データフレームから条件指定で行データを抽出する
        DF = df[df[label].str.contains(title)]
        
        # プロットするカラーを指定するための処理
        cnt = len(DF.groupby(label).count().index.to_list())
        color_list = []
        for j in range(cnt):
            color_list.append(cm.hsv(j/cnt)) # jet cool autumn hsv
        color_list

        # ひとつのグラフにカテゴリ別にプロットする
        handle_list = []
        for i, legend in enumerate(DF[label].unique()):
            # ラベルを文字列から絞って抽出する
            index = legend.index("_")
            label_name = legend[index + 1:]
            
            # 折れ線図を作成
            plt.plot(DF.loc[DF[label] == legend, x_name],
                        DF.loc[DF[label] == legend, y_name],
                        #facecolor = 'None',
                        #edgecolors = color_list[i],
                        color = color_list[i],
                        label = label_name)
        plt.grid()
        plt.legend(bbox_to_anchor = (1, 1), title = legend_name)
        plt.xlabel(x_name)
        plt.ylabel(y_name)
        plt.title(title)

        # 保存するファイルパス
        save_jpg_path = f"{self.folder_path}/{self.file_name}.jpg"
        plt.savefig(save_jpg_path, bbox_inches="tight")
        print(f"save {save_jpg_path}")
        print('finished')

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

