import sys
from PySide6.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QComboBox, QDateEdit, QPushButton
from PySide6.QtCore import QDate

class ReportProcessor(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        input_layouts = [
            ("姓名", QLineEdit()),
            ("性别", QComboBox()),
            ("年龄", QLineEdit()),
            ("标本类型", QComboBox()),
            ("送检日期", QDateEdit()),
            ("科室", QLineEdit()),  # 改为 QLineEdit，支持自定义输入
            ("门诊号/住院号", QLineEdit()),
            ("临床诊断", QLineEdit()),
            ("送检医生", QComboBox())
        ]

        for label_text, widget in input_layouts:
            hlayout = QHBoxLayout()
            label = QLabel(label_text)
            hlayout.addWidget(label)
            hlayout.addWidget(widget)
            widget.setFixedHeight(40)
            layout.addLayout(hlayout)

        
        self.user_name = input_layouts[0][1]
        
        self.gender_combo = input_layouts[1][1]
        self.gender_combo.addItems(["男", "女"])

        self.user_age = input_layouts[2][1]

        self.specimen_combo = input_layouts[3][1]
        self.specimen_combo.addItems(["血液", "尿液", "组织样本"])

        self.date_edit = input_layouts[4][1]
        self.date_edit.setDate(QDate.currentDate())

        # 送检医生的下拉菜单初始为空
        self.doctor_combo = input_layouts[8][1]


        self.build_file = QPushButton("生成报告")
        layout.addWidget(self.build_file)        
        self.build_file.clicked.connect(self.build_file_handle)
        
        
        self.setLayout(layout)

        self.setWindowTitle("检测报告处理")
        
        self.setFixedWidth(600)
        
    def build_file_handle(self):
        # print(dir(self.user_name))
        print(self.user_name.text(), self.user_age.text())
        
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ReportProcessor()
    window.show()
    sys.exit(app.exec())