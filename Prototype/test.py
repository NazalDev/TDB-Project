from PyQt5.QtWidgets import QApplication, QMainWindow, QComboBox, QLineEdit, QLabel
import sys

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Auto-Fill Form Example")
        self.setGeometry(100, 100, 300, 200)

        # ComboBox
        self.combo = QComboBox(self)
        self.combo.setGeometry(50, 30, 200, 30)
        self.combo.addItems(["Select Product", "Apple", "Banana", "Cherry"])

        # LineEdits for auto-filled data
        self.price_field = QLineEdit(self)
        self.price_field.setGeometry(50, 70, 200, 30)
        self.price_field.setPlaceholderText("Price")

        self.stock_field = QLineEdit(self)
        self.stock_field.setGeometry(50, 110, 200, 30)
        self.stock_field.setPlaceholderText("Stock")

        # Connect combo box selection to handler
        self.combo.currentTextChanged.connect(self.auto_fill_fields)

    def auto_fill_fields(self, text):
        # Simulated data lookup
        product_data = {
            "Apple": {"price": "10", "stock": "50"},
            "Banana": {"price": "5", "stock": "100"},
            "Cherry": {"price": "15", "stock": "30"}
        }

        if text in product_data:
            self.price_field.setText(product_data[text]["price"])
            self.stock_field.setText(product_data[text]["stock"])
        else:
            self.price_field.clear()
            self.stock_field.clear()

app = QApplication(sys.argv)
window = MainWindow()
window.show()
sys.exit(app.exec_())