## This file is for button only not all widgets


from PyQt5 import QtCore, QtGui, QtWidgets

class UI_Buttons(object):
    def UI_Buttons(self, MainWindow):
        ## view more button for table in home page
        self.view_more = QtWidgets.QPushButton(MainWindow)
        self.view_more.setText("View More")
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/Green/assets/icon/green/zoom-in.svg"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        icon.addPixmap(QtGui.QPixmap(":/Green/assets/icon/white/zoom-in.svg"), QtGui.QIcon.Normal, QtGui.QIcon.On)
        self.view_more.setIcon(icon)
        self.view_more.setObjectName("view_more")