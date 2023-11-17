import sys
import csv

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QImage, QPixmap
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QLineEdit, QTextEdit, QPushButton, QWidget, \
    QVBoxLayout, QFileDialog, QDialog, QHBoxLayout, QMessageBox
from docx_convertor import MyDocument


class DataEntryApp(QMainWindow, MyDocument):
    def __init__(self):
        super().__init__()
        self.file_names = [0, 0]
        self.setFixedWidth(900)
        self.setWindowTitle("Lab2")
        self.setStyleSheet('''
                            QMainWindow{
                                background-colot: gray;
                                color: black;
                            }
                            QPushButton {
                                background-color: lightgray;
                                border-style: solid;
                                border-width: 2px;
                                border-color: black;
                                font: 12px;
                                min-width: 8em;
                                padding: 6px;й  
                            }
                            QPushButton:pressed {
                                background-color: darkgray;
                                border-style: inset;
                            }
                            ''')
        self.initUI()

    def initUI(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout()
        layout = QHBoxLayout()
        labels = [
            "Полное название организации",
            "Название кафедры",
            "Номер лабораторной",
            "Название дисциплины",
            "Номер группы",
            "Фамилия И.О. студента",
            "Фамилия И.О. проверяющего",
            "Год написания",
            "Номер задания",
            "Язык програмированния",
            "Цель работы",
            "Текст задания",
            "Решение",
            "Листинг программы",
            "Вывод",
            "Импортируйте картинку 1",
            "Подпись к картинке 1",
            "Импортируйте картинку 2",
            "Подпись к картинке 2"
        ]

        self.input_fields = []
        row_layout = QVBoxLayout()
        for i in range(0, 10):
            label = QLabel(labels[i])
            label.setMaximumWidth(160)
            input_field = QLineEdit()
            input_field.setFixedWidth(160)

            row_layout.addWidget(label)
            row_layout.addWidget(input_field)

            self.input_fields.extend([input_field])

        layout.addLayout(row_layout)
        fields_layout = QVBoxLayout()
        space = QLabel()
        space.setFixedWidth(20)

        for i in range(10, len(labels)):
            row_layout = QHBoxLayout()

            label = QLabel(labels[i])
            label.setFixedWidth(160)
            browse_button = QPushButton()
            if 15 <= i <= 18 and i % 2 != 0:
                input_field = QLineEdit()
                input_field.setEnabled(False)
                input_field.setFixedWidth(160)
                browse_button = QPushButton("Добавить изображение")
                browse_button.setFixedWidth(335)
                browse_button.clicked.connect(lambda checked, index=i: self.browse_image(index))

            elif i >= 15:
                input_field = QLineEdit()
                input_field.setFixedWidth(500)
            else:
                input_field = QTextEdit()
                input_field.setFixedSize(500, 80)

            row_layout.addWidget(label)
            row_layout.addWidget(input_field)
            if i == 15 or i == 17:
                row_layout.addWidget(browse_button)
            fields_layout.addLayout(row_layout)
            self.input_fields.append(input_field)
        fields_layout.setAlignment(Qt.AlignRight)
        layout.addLayout(fields_layout)
        main_layout.addLayout(layout)
        save_button = QPushButton("Сохранить файл")
        save_button.clicked.connect(self.show_save_dialog)
        main_layout.addWidget(save_button)

        central_widget.setLayout(main_layout)

    def browse_image(self, index):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_name, _ = QFileDialog.getOpenFileName(self, 'QFileDialog.getOpenFileName()', '',
                                                   'All Files (*);;Image Files (*.png *.jpg *.jpeg *.bmp)',
                                                   options=options)
        if file_name:
            self.input_fields[index].setText(file_name)
            self.file_names[(index - 15) // 2] = file_name

    def show_save_dialog(self):
        dialog = QDialog()
        dialog.setWindowTitle("Сохранение документа")
        vbox = QVBoxLayout()
        image_label = QLabel("Выбранные изображения")
        image_box = QHBoxLayout()
        image_labels = [ImageLabel(self) for i in range(2)]
        for i in image_labels:
            i.setStyleSheet("border-style: solid; border-width: 1px; border-color: black;")
            file_name = self.file_names[image_labels.index(i)]
            if file_name != 0:
                i.set_image(file_name)
            i.setFixedSize(300, 300)
            image_box.addWidget(i)

        vbox.addWidget(image_label)
        vbox.addLayout(image_box)
        hbox = QHBoxLayout()
        label = QLabel("Желаете сохранить документ?")
        hbox.addWidget(label)

        save_button = QPushButton("Сохранить документ")
        save_button.clicked.connect(self.save_document)
        hbox.addWidget(save_button)

        close_button = QPushButton("Закрыть")
        close_button.clicked.connect(dialog.accept)
        hbox.addWidget(close_button)
        vbox.addLayout(hbox)

        dialog.setLayout(vbox)
        dialog.exec()

    def save_document(self):
        if self.validate_user_data():
            user_data = self.get_user_data()
            for key, value in user_data.items():
                print(f"{key}: {value}")

            with open('user_data.csv', 'w', newline='') as csvfile:
                csv_writer = csv.writer(csvfile)
                csv_writer.writerow(user_data.keys())
                csv_writer.writerow(user_data.values())
            self.create_lab_report()
            QMessageBox.information(self, "Документ сохранен", "Документ сохранен успешно.")
        else:
            QMessageBox.warning(self, "Не возможно сохранить документ",
                                "Для сохранения документа нужно заполнить все поля.")

    def validate_user_data(self):
        for field in self.input_fields:
            if isinstance(field, QLineEdit):
                if field.text().strip() == "":
                    return False
            elif isinstance(field, QTextEdit):
                if field.toPlainText().strip() == "":
                    return False
        return True

    def get_user_data(self):
        user_data = {}

        for i, label in enumerate(self.input_fields):
            user_data[f"Field {i + 1}"] = label.text() if isinstance(label, QLineEdit) else label.toPlainText()

        return user_data


class ImageLabel(QLabel):
    def __init__(self, parent=None):
        super().__init__(parent)

    def set_image(self, image):
        self.image = QImage(image).scaled(300, 300)
        self.setPixmap(QPixmap.fromImage(self.image))


QSS = '''
QPushButton {
    background-color: darkgray;
    border-style: solid;
    border-width: 2px;
    border-color: black;
    font: 12px;
    min-width: 8em;
    padding: 6px;
}
QPushButton:pressed {
    background-color: lightgray;
    border-style: inset;
}
'''

if __name__ == '__main__':
    app = QApplication([])
    window = DataEntryApp()
    window.show()

    sys.exit(app.exec())
