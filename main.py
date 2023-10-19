#Main function which will create PySide Application and run it. The Application will be opened full screen and will have dark theme style.

from PySide6.QtWidgets import QApplication
from lib.ui.mainwindow import MainWindow

app = QApplication([])
window = MainWindow()
window.show()
app.exec()