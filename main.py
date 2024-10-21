import sys, PyPDF2, os, reportlab, PyQt5
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
from PyQt5.uic import loadUi
from PyQt5.QtCore import QSettings
from datetime import date
from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.pagesizes import landscape, A4, letter
from PyPDF2 import PdfWriter, PdfReader, PdfMerger
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics


class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        loadUi('mergeDate.ui', self)
        text_in_edit = date.today().strftime("%d.%m.%Y")
        self.settings = QSettings("MyApp", "MyProgram")
        self.textEdit.setText(text_in_edit)
        self.textEdit_2.setText("Куцин О.В.")
        self.textEdit_3.setText("Бенца Р.А.")
        self.pushButton.clicked.connect(self.add_date_to_pdf)
        self.pushButton_2.clicked.connect(self.merge_files)
        self.pushButton_3.clicked.connect(self.delete_dates_from_pdf)
        self.pushButton_4.clicked.connect(self.add_names)
        pdfmetrics.registerFont(TTFont('myArial', 'arial.ttf'))

    def add_date_to_pdf(self):
        if self.textEdit.toPlainText() == "":
            error = QMessageBox()
            error.setWindowTitle("Помилка!")
            error.setText("Введіть дату!!!")
            error.setIcon(QMessageBox.Information)
            error.exec_()
        else:
            file_dialog = QFileDialog()
            file_dialog.setFileMode(QFileDialog.ExistingFiles)
            last_used_path = self.settings.value("last_used_path")
            if last_used_path:
                file_dialog.setDirectory(last_used_path)
            else:
                file_dialog.setDirectory("D:\\")
            file_paths, _ = file_dialog.getOpenFileNames(self, "Оберіть файл(и)", "", "PDF Files (*.pdf)")
            if file_paths:
                last_used_path = file_paths[0]
                self.settings.setValue("last_used_path", last_used_path)
                flag = False
                canvas = Canvas("portrait.pdf")
                canvas.setFont("myArial", 11)
                canvas.setFillColorRGB(1, 1, 1)
                canvas.rect(315, 20, 80, 20, fill=True, stroke=False)
                canvas.setFillColorRGB(0, 0, 0)
                canvas.drawString(325, 25, self.textEdit.toPlainText())
                canvas.save()

                canvas1 = Canvas("landscape.pdf")
                width, height = A4
                canvas1.setPageSize((height, width))
                canvas1.setFont("myArial", 11)
                canvas1.setFillColorRGB(1, 1, 1)
                canvas1.rect(560, 20, 80, 20, fill=True, stroke=False)
                canvas1.setFillColorRGB(0, 0, 0)
                canvas1.drawString(570, 25, self.textEdit.toPlainText())
                canvas1.save()

                for paths in file_paths:
                    writer = PdfWriter()
                    reader = PdfReader(paths)
                    portrait = PdfReader("portrait.pdf")
                    landscape = PdfReader("landscape.pdf")
                    for pages in reader.pages:
                        if pages.mediabox.width > pages.mediabox.height:
                            pages.merge_page(landscape.pages[0])
                            writer.add_page(pages)
                        else:
                            pages.merge_page(portrait.pages[0])
                            writer.add_page(pages)
                    os.remove(paths)
                    with open(os.path.dirname(file_paths[0]) + "\\" + os.path.basename(paths), "wb") as fp:
                        writer.write(fp)
                        flag = True
                if flag:
                    success = QMessageBox()
                    success.setWindowTitle("Успіх!")
                    success.setText("Дати проставлені успішно")
                    success.setIcon(QMessageBox.Information)
                    success.exec_()
                    os.remove("portrait.pdf")
                    os.remove("landscape.pdf")
            else:
                error = QMessageBox()
                error.setWindowTitle("Помилка!")
                error.setText("Оберіть файл або декілька файлів!")
                error.setIcon(QMessageBox.Information)
                error.exec_()
    def merge_files(self):
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFiles)
        last_used_path = self.settings.value("last_used_path")
        if last_used_path:
            file_dialog.setDirectory(last_used_path)
        else:
            file_dialog.setDirectory("D:\\")

        file_paths, _ = file_dialog.getOpenFileNames(self, "Оберіть файл(и)", "", "PDF Files (*.pdf)")
        if len(file_paths) < 2:
            error = QMessageBox()
            error.setWindowTitle("Помилка!")
            error.setText("Оберіть більше файлів!")
            error.setIcon(QMessageBox.Information)
            error.exec_()
        else:
            last_used_path = file_paths[0]
            self.settings.setValue("last_used_path", last_used_path)
            merger = PdfMerger()
            for file_path in file_paths:
                with open(file_path, "rb") as file:
                    merger.append(file)
            merger.write(os.path.dirname(file_paths[0]) + "/merged.pdf")
            merger.close()
            success = QMessageBox()
            success.setWindowTitle("Успіх!")
            success.setText("Файли об'єднано успішно")
            success.setIcon(QMessageBox.Information)
            success.exec_()

    def delete_dates_from_pdf(self):
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFiles)
        last_used_path = self.settings.value("last_used_path")
        if last_used_path:
            file_dialog.setDirectory(last_used_path)
        else:
            file_dialog.setDirectory("С:\\")
        file_paths, _ = file_dialog.getOpenFileNames(self, "Оберіть файл(и)", "", "PDF Files (*.pdf)")
        if file_paths:
            last_used_path = file_paths[0]
            self.settings.setValue("last_used_path", last_used_path)
            flag = False
            canvas = Canvas("portrait.pdf")
            canvas.setFillColorRGB(1, 1, 1)
            canvas.rect(315, 20, 80, 20, fill=True, stroke=False)
            canvas.rect(308, 55, 80, 14, fill=True, stroke=False)
            canvas.rect(489, 55, 80, 14, fill=True, stroke=False)
            canvas.save()

            canvas1 = Canvas("landscape.pdf")
            width, height = A4
            canvas1.setPageSize((height, width))
            canvas1.setFillColorRGB(1, 1, 1)
            canvas1.rect(560, 20, 80, 20, fill=True, stroke=False)
            canvas1.rect(555, 55, 80, 14, fill=True, stroke=False)
            canvas1.rect(735, 55, 80, 14, fill=True, stroke=False)
            canvas1.save()
            for paths in file_paths:
                writer = PdfWriter()
                reader = PdfReader(paths)
                portrait = PdfReader("portrait.pdf")
                landscape = PdfReader("landscape.pdf")
                for pages in reader.pages:
                    if pages.mediabox.width > pages.mediabox.height:
                        pages.merge_page(landscape.pages[0])
                        writer.add_page(pages)
                    else:
                        pages.merge_page(portrait.pages[0])
                        writer.add_page(pages)
                os.remove(paths)
                with open(os.path.dirname(file_paths[0]) + "\\" + os.path.basename(paths), "wb") as fp:
                    writer.write(fp)
                    flag = True
            if flag:
                success = QMessageBox()
                success.setWindowTitle("Успіх!")
                success.setText("Дати видалені успішно")
                success.setIcon(QMessageBox.Information)
                success.exec_()
                os.remove("portrait.pdf")
                os.remove("landscape.pdf")
            else:
                error = QMessageBox()
                error.setWindowTitle("Помилка!")
                error.setText("Оберіть файл або декілька файлів!")
                error.setIcon(QMessageBox.Information)
                error.exec_()

    def add_names(self):
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFiles)
        last_used_path = self.settings.value("last_used_path")
        if last_used_path:
            file_dialog.setDirectory(last_used_path)
        else:
            file_dialog.setDirectory("С:\\")
        file_paths, _ = file_dialog.getOpenFileNames(self, "Оберіть файл(и)", "", "PDF Files (*.pdf)")
        if file_paths:
            last_used_path = file_paths[0]
            self.settings.setValue("last_used_path", last_used_path)
            flag = False
            canvas = Canvas("portrait.pdf")
            canvas.setFillColorRGB(1, 1, 1)
            canvas.rect(308, 55, 80, 14, fill=True, stroke=False)
            canvas.rect(489, 55, 80, 14, fill=True, stroke=False)
            canvas.setFont("myArial", 10)
            canvas.setFillColorRGB(0, 0, 0)
            canvas.drawString(313, 60, self.textEdit_2.toPlainText())
            canvas.drawString(494, 60, self.textEdit_3.toPlainText())
            canvas.save()

            canvas1 = Canvas("landscape.pdf")
            width, height = A4
            canvas1.setPageSize((height, width))
            canvas1.setFillColorRGB(1, 1, 1)
            canvas1.rect(555, 55, 80, 14, fill=True, stroke=False)
            canvas1.rect(735, 55, 80, 14, fill=True, stroke=False)
            canvas1.setFont("myArial", 10)
            canvas1.setFillColorRGB(0, 0, 0)
            canvas1.drawString(560, 60, self.textEdit_2.toPlainText())
            canvas1.drawString(740, 60, self.textEdit_3.toPlainText())
            canvas1.save()
            for paths in file_paths:
                writer = PdfWriter()
                reader = PdfReader(paths)
                portrait = PdfReader("portrait.pdf")
                landscape = PdfReader("landscape.pdf")
                for pages in reader.pages:
                    if pages.mediabox.width > pages.mediabox.height:
                        pages.merge_page(landscape.pages[0])
                        writer.add_page(pages)
                    else:
                        pages.merge_page(portrait.pages[0])
                        writer.add_page(pages)
                os.remove(paths)
                with open(os.path.dirname(file_paths[0]) + "\\" + os.path.basename(paths), "wb") as fp:
                    writer.write(fp)
                    flag = True
            if flag:
                success = QMessageBox()
                success.setWindowTitle("Успіх!")
                success.setText("Прізвища проставлені успішно")
                success.setIcon(QMessageBox.Information)
                success.exec_()
                os.remove("portrait.pdf")
                os.remove("landscape.pdf")
            else:
                error = QMessageBox()
                error.setWindowTitle("Помилка!")
                error.setText("Оберіть файл або декілька файлів!")
                error.setIcon(QMessageBox.Information)
                error.exec_()


app = QApplication(sys.argv)
window = MainWindow()
window.show()
sys.exit(app.exec_())
