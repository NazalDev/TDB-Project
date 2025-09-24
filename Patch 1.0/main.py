import sys

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QMainWindow, QApplication, QPushButton, QGraphicsDropShadowEffect, QWidget
from PyQt5.QtCore import pyqtSlot, QFile, QTextStream, QEvent, QPropertyAnimation, QRect
from PyQt5.QtGui import QColor


from mainWindow_ui import Ui_MainWindow
from poForm import PoForm;
import customButton


class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()

        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        self.ui.full_sidebar.hide()
        self.ui.stackedWidget.setCurrentIndex(1)
        self.ui.home_icon.setChecked(True)
        
        ## Clicked Button function
        self.ui.po_tab_add_po_btn.clicked.connect(self.poFormConnect)

        ## hovering system
        self.ui.icon_sidebar.installEventFilter(self)
        self.ui.full_sidebar.installEventFilter(self)
        
        ## Adding shadow
        shadowColor = QColor(0, 0, 0, 125)

        self.ui.icon_sidebar.setGraphicsEffect(self.dropShadow(24, 3, 0, shadowColor))
        self.ui.full_sidebar.setGraphicsEffect(self.dropShadow(24, 3, 0, shadowColor))
        self.ui.poDo_tabWidget.tabBar().setGraphicsEffect(self.dropShadow(20,0,7, shadowColor))

        self.ui.header.setGraphicsEffect(self.dropShadow(24, 0, 3))

        ## Load Data
        # Load Po Table
        self.loadPoTable()

        ## self explanatory function dump


    ##################################################################################################################################################################
    ## CHECKABLE BUTTON CHANGE WHEN STACKEDWIDGET INDEX CHANGED START
    
    def on_stackedWidget_currentChanged(self, index):
        btn_list = self.ui.icon_sidebar.findChildren(QPushButton) \
                    + self.ui.full_sidebar.findChildren(QPushButton)
        
        for btn in btn_list:
            if  index in [0]:
                btn.setAutoExclusive(False)
                btn.setChecked(False)
            else:
                btn.setAutoExclusive(True)

    # Changing menu page
    
    def on_home_full_toggled(self):
        self.ui.stackedWidget.setCurrentIndex(1)
    def on_home_icon_toggled(self):
        self.ui.stackedWidget.setCurrentIndex(1)

    def on_dashboard_full_toggled(self):
        self.ui.stackedWidget.setCurrentIndex(2)
    def on_dashboard_icon_toggled(self):
        self.ui.stackedWidget.setCurrentIndex(2)

    def on_report_full_toggled(self):
        self.ui.stackedWidget.setCurrentIndex(3)
    def on_report_icon_toggled(self):
        self.ui.stackedWidget.setCurrentIndex(3)

    def on_poDo_full_toggled(self):
        self.ui.stackedWidget.setCurrentIndex(4)
    def on_poDo_icon_toggled(self):
        self.ui.stackedWidget.setCurrentIndex(4)

    def on_product_full_toggled(self):
        self.ui.stackedWidget.setCurrentIndex(5)
    def on_product_icon_toggled(self):
        self.ui.stackedWidget.setCurrentIndex(5)

    def on_invoice_full_toggled(self):
        self.ui.stackedWidget.setCurrentIndex(6)
    def on_invoice_icon_toggled(self):
        self.ui.stackedWidget.setCurrentIndex(6)
                

    ## CHECKABLE BUTTON CHANGE WHEN STACKEDWIDGET INDEX CHANGED END
    ##################################################################################################################################################################
    ##  PO Table START

    def loadPoTable(self):
        self.ui.po_table.setHorizontalHeaderLabels(["", "", "Purchase No.", "Purchase Date", "Perusahaan", "Site", "Produk"])

        self.ui.po_table.clearContents()
        self.ui.po_table.hideColumn(0)
        self.ui.po_table.hideColumn(1)

        ## View More button Styling
        ##for row in result:
        ##   viewMoreBtn = customButton.viewMoreButton()
        ##    self.ui.po_table.setCellWidget(0, 0, viewMoreBtn)
        
        

    ##  PO Table END
    ##################################################################################################################################################################
    ## ANIMATION WIDGET START

    def animationSidebar(self):
        main_page_width = self.ui.main_page.width()
        main_page_height = self.ui.main_page.height()
        self.main_page_animation = QPropertyAnimation(self.ui.main_page, b"geometry")
        self.main_page_animation.setDuration(300) # mm seconds
        self.main_page_animation.setStartValue(QRect(self.ui.main_page.x(), self.ui.main_page.y(), main_page_width, main_page_height))
        self.main_page_animation.setEndValue(QRect(self.ui.main_page.x() + 174, self.ui.main_page.y(), main_page_width - 174, main_page_height))
        self.main_page_animation.start()
        
        sidebar_height = self.height()
        self.ui.full_sidebar.show()
        self.sidebar = QPropertyAnimation(self.ui.full_sidebar, b"geometry")
        self.sidebar.setDuration(300) # mm seconds
        self.sidebar.setStartValue(QRect(self.ui.full_sidebar.x(), self.ui.full_sidebar.y(), 0, sidebar_height))
        self.sidebar.setEndValue(QRect(self.ui.full_sidebar.x(), self.ui.full_sidebar.y(), 240, sidebar_height))
        self.sidebar.start()
        


    ## ANIMATION WIDGET END
    ##################################################################################################################################################################
    ## Function Dump

    ## Applying shadow function
    def dropShadow(self, blurRadius : float, xOffset : int, yOffset : int, color = None):
        shadow = QGraphicsDropShadowEffect()
        if blurRadius is not None:
            shadow.setBlurRadius(blurRadius)
        shadow.setOffset(xOffset if xOffset is not None else 0, yOffset if yOffset is not None else 0)
        if color is not None:
            shadow.setColor(color)
        
        return shadow


    ## Cursor hover identification function
    def eventFilter(self, obj, event):
        if obj == self.ui.icon_sidebar:
            if event.type() == QEvent.Enter:  # cursor enters button
                self.ui.icon_sidebar.hide()
                self.animationSidebar()

        elif obj == self.ui.full_sidebar:
            if event.type() == QEvent.Leave:  # cursor leaves button
                self.ui.full_sidebar.hide()
                self.ui.icon_sidebar.show()
        return super().eventFilter(obj, event)

    ## Function Dump End
    ##################################################################################################################################################################
    ## Connecting Function Start
    def poFormConnect(self):
        self.form = PoForm()
        self.form.show()


    ## Connecting Function End
    ##################################################################################################################################################################



if __name__ == "__main__":
    app = QApplication(sys.argv)

    ## loading style file
    with open("style.qss", "r") as style_file:
        style_str = style_file.read()
    app.setStyleSheet(style_str)

    window = MainWindow()
    window.show()

    sys.exit(app.exec())