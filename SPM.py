import os
import subprocess
import sys
from os.path import expanduser
from pathlib import Path
import shutil

import docx
import nltk
import string
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QStandardItemModel
from PyQt5.QtWidgets import *
from sklearn.feature_extraction.text import TfidfVectorizer


class App(QWidget):
    FILES, SIMILAR, FP = range(3)

    def __init__(self):
        super().__init__()

        self.vectorizer = TfidfVectorizer(tokenizer=self.normalize, stop_words='english')

        self.title = 'String Pattern Matcher'
        self.left = 50
        self.top = 50
        self.width = 840
        self.height = 340
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        self.dataGroupBox = QGroupBox("Similar Files")
        self.dataView = QTreeView()
        self.dataView.setRootIsDecorated(False)
        self.dataView.setAlternatingRowColors(True)

        self.dataView.setSortingEnabled(True)
        self.dataView.clicked.connect(self.on_clicked)

        self.dataView.setContextMenuPolicy(Qt.CustomContextMenu)
        self.dataView.customContextMenuRequested.connect(self.tabMenu)

        self.dataBtn = QPushButton('Check')
        self.dataBtn.setFixedSize(75, 30)
        self.dataBtn.clicked.connect(self.check)

        self.notify = QLabel('Notify')
        self.notify.setFixedSize(75, 30)

        self.addBtn = QPushButton('Add New File')
        self.addBtn.setFixedSize(75, 30)
        self.addBtn.clicked.connect(self.copyfile)

        self.glayout = QGridLayout()
        self.glayout.addWidget(self.dataBtn, 0, 0)
        self.glayout.addWidget(self.notify, 0, 1)
        self.glayout.addWidget(self.addBtn, 0, 2)

        dataLayout = QVBoxLayout()
        dataLayout.addWidget(self.dataView)
        dataLayout.addLayout(self.glayout)

        self.dataGroupBox.setLayout(dataLayout)

        mainLayout = QVBoxLayout()
        mainLayout.addWidget(self.dataGroupBox)

        self.setLayout(mainLayout)

        self.show()

    def createFileModel(self, parent):
        model = QStandardItemModel(0, 3, parent)
        model.setHeaderData(self.FILES, Qt.Horizontal, " Files")
        model.setHeaderData(self.SIMILAR, Qt.Horizontal, " %  Similarities")
        model.setHeaderData(self.FP, Qt.Horizontal, " File Path")
        return model

    def addFile(self, model, files, similar, file_path):
        model.insertRow(0)
        model.setData(model.index(0, self.FILES), files)
        model.setData(model.index(0, self.SIMILAR), similar)
        model.setData(model.index(0, self.FP), file_path)

    def getText(self, filename):
        try:
            fx = filename.split(".")[-1]
            fullText = []
            if fx == 'docx':
                doc = docx.Document(filename)
                for para in doc.paragraphs:
                    fullText.append(para.text)
            elif fx == 'doc' or 'txt':
                file = open(filename, "r", encoding="ascii", errors="ignore").read()
                fullText.append(file)
            else:
                QMessageBox.information(self, 'Error!', 'Not A Document File!')
                return False
            return '\n'.join(fullText)
        except Exception:
            QMessageBox.information(self, 'Error!', 'Not A Document File!')

    # nltk.download('punkt')  # if necessary...

    stemmer = nltk.stem.porter.PorterStemmer()
    remove_punctuation_map = dict((ord(char), None) for char in string.punctuation)

    def stem_tokens(self, tokens):
        return [self.stemmer.stem(item) for item in tokens]

    '''remove punctuation, lowercase, stem'''

    def normalize(self, text):
        return self.stem_tokens(nltk.word_tokenize(text.lower().translate(self.remove_punctuation_map)))

    def cosine_sim(self, text1, text2):
        text11 = text1
        text22 = open(text2, 'r', encoding='utf-8', errors='ignore').read()
        tfidf = self.vectorizer.fit_transform([text11, text22])

        n = (((tfidf * tfidf.T) * 100).A)[0, 1]
        return '%.3f%% ' %n

    def check(self):
        self.notify.show()
        self.model = self.createFileModel(self)
        self.dataView.setModel(self.model)
        self.notify.setText('Checking...')

        try:
            fpath = QFileDialog.getOpenFileName(self, 'Select File to check',
                                                expanduser("~"))
            self.notify.setText('Checking...')

            self.file = str(fpath[0]).replace('/', '\\')

            print(self.file)

            spath = expanduser('~') + '\Desktop\Work'

            text = self.getText(self.file)
            if os.path.exists(spath):
                for path in Path(spath).iterdir():
                    self.addFile(self.model, os.path.basename(path), self.cosine_sim(text, path), str(path))
            self.notify.setText('Done')
        except Exception:
            self.notify.hide()
            pass

    def on_clicked(self):

        try:
            for ix in self.dataView.selectedIndexes():
                if ix.column() == 2:
                    self.path = ix.data()
                    print(self.path)

        except Exception:
            print('Index error')

    def tabMenu(self, position):

        self.tmenu = QMenu()

        self.open = self.tmenu.addAction('Open')
        self.open_file_location = self.tmenu.addAction('Open File Location')

        self.tmenu.addActions([self.open, self.open_file_location])
        action = self.tmenu.exec_(self.dataView.viewport().mapToGlobal(position))

        if action == self.open:
            os.startfile(self.path, 'open')
        elif action == self.open_file_location:
            try:
                subprocess.Popen(r'explorer /select,' + "{}".format(self.path).replace('/', '\\'))
            except Exception:
                pass

    def copyfile(self):
        try:
            if os.path.exists(self.file):
                src = os.path.realpath(self.file)
                nsrc = os.path.basename(src)

                dpath = expanduser('~') + '\\Desktop\\Work\\'

                dst = dpath + nsrc + '.copy'

                shutil.copy(src, dst)
                self.notify.setText('New Doc Added')
                QMessageBox.information(self, 'Done !', 'New Document Added!')
            else:
                print('path error')
        except Exception:
            QMessageBox.information(self, 'Error', 'Select Document First!')

    def closeEvent(self, event):
        reply = QMessageBox.question(self, 'Exit', "Are you sure you want to exit?", QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())
