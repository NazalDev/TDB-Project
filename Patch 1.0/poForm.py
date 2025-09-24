from PyQt5.QtWidgets import QWidget

from poForm_ui import Ui_Form as poForm

class PoForm(QWidget):
    def __init__(self):
        super(PoForm, self).__init__()

        self.ui = poForm()
        self.ui.setupUi(self)

        self.ui.companyInfo_widget.hide()
        self.ui.siteInfo_widget.hide()
        self.ui.company_combo.addItem("TDB")

        ## Auto update date combo box
        for i in range(1, 32):
            self.ui.poDate_day_combo.addItem(str(i))

        self.ui.company_combo.currentTextChanged.connect(self.getText)
        self.ui.site_combo.currentTextChanged.connect(self.getText)
        self.ui.poDate_month_combo.currentTextChanged.connect(self.dateAdd)

    def getText(self, text):
        text1 = self.ui.company_combo.currentText()
        text = self.ui.site_combo.currentText()

        if text1 == "New company...":
            self.ui.companyInfo_widget.show()
        else:
            self.ui.companyInfo_widget.hide()

        if text == "New site...":
            self.ui.siteInfo_widget.show()
        else:
            self.ui.siteInfo_widget.hide()

    def dateAdd(self, text):
        maxDate = 31

        match text:
            case "Januari":
                maxDate = 31
            case "Februari":
                maxDate = 29
            case "Maret":
                maxDate = 31
            case "April":
                maxDate = 30
            case "Mei":
                maxDate = 31
            case "Juni":
                maxDate = 30
            case "Juli":
                maxDate = 31
            case "Agustus":
                maxDate = 31
            case "September":
                maxDate = 30
            case "Oktober":
                maxDate = 31
            case "November":
                maxDate = 30
            case "Desember":
                maxDate = 31

        for i in range(1, maxDate+1):
            self.ui.poDate_day_combo.addItem(str(i))

        

