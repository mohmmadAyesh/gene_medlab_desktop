# login.py
import bcrypt
from PySide6.QtWidgets import QDialog, QFormLayout, QLineEdit, QPushButton, QMessageBox
from models import User

class LoginDialog(QDialog):
    def __init__(self, db_session):
        super().__init__()
        self.db = db_session
        self.current_user = None
        self.setWindowTitle("Login")
        form = QFormLayout(self)
        self.user_edit = QLineEdit(); form.addRow("Username:", self.user_edit)
        self.pw_edit   = QLineEdit(); self.pw_edit.setEchoMode(QLineEdit.Password)
        form.addRow("Password:", self.pw_edit)
        btn_ok    = QPushButton("OK");     btn_ok.clicked.connect(self.attempt)
        btn_canc  = QPushButton("Cancel"); btn_canc.clicked.connect(self.reject)
        form.addRow(btn_ok, btn_canc)

    def attempt(self):
        user = self.db.query(User).filter_by(username=self.user_edit.text()).first()
        pw   = self.pw_edit.text().encode()
        if user and bcrypt.checkpw(pw, user.password_hash.encode()):
            self.current_user = user
            self.accept()
        else:
            QMessageBox.warning(self, "Login Failed", "Invalid credentials")
