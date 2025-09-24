import resources_rc

from PyQt5 import QtGui
from PyQt5.QtWidgets import QPushButton, QCheckBox

def viewMoreButton() -> QPushButton:
    viewMoreBtn = QPushButton()
    icon1 = QtGui.QIcon()
    icon1.addPixmap(QtGui.QPixmap(":/assets/assets/icon/dark_green/search.svg"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
    viewMoreBtn.setIcon(icon1)

    return viewMoreBtn

def checkBox() -> QCheckBox:
    deleteCheckBtn = QCheckBox()
    icon1 = QtGui.QIcon()
    icon1.addPixmap(QtGui.QPixmap(":/assets/assets/icon/dark_green/checkbox.svg"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
    deleteCheckBtn.setIcon(icon1)

    return deleteCheckBtn