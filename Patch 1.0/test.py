import sys
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QPushButton, QVBoxLayout, QLineEdit

class EditableLabel(QLabel):
    def __init__(self, text):
        super().__init__(text)
        self.setText(text)
        self.setStyleSheet("background: #FFF9E5; border: 1px solid #ccc; padding: 4px;")
        self.setCursor(Qt.IBeamCursor)
        self.edit = None

    def mouseDoubleClickEvent(self, event):
        if not self.edit:
            self.edit = QLineEdit(self.text(), self.parent())
            self.edit.setGeometry(self.geometry())
            self.edit.show()
            self.edit.setFocus()
            self.edit.editingFinished.connect(self.finishEdit)
            self.hide()

    def finishEdit(self):
        self.setText(self.edit.text())
        self.show()
        self.edit.deleteLater()
        self.edit = None

class DupeDemo(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Duplicate widget example")
        self.setGeometry(100, 100, 300, 300)

        self.layout = QVBoxLayout(self)
        self.setLayout(self.layout)

        self.duplicate_btn = QPushButton("Duplicate Label")
        self.duplicate_btn.clicked.connect(self.duplicate_label)
        self.layout.addWidget(self.duplicate_btn)

        self.label = EditableLabel("Double-click to edit me!")
        self.layout.addWidget(self.label)

    def duplicate_label(self):
        new_label = EditableLabel(self.label.text())
        self.layout.addWidget(new_label)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = DupeDemo()
    win.show()
    sys.exit(app.exec_())