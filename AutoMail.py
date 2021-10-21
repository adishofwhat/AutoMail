from PyQt5 import QtWidgets, QtGui
from PyQt5.QtWidgets import QApplication, QMainWindow, QDialog, QFileDialog, QLabel, QComboBox, QLineEdit
from PyQt5.QtGui import QFont
import sys
from PyQt5.uic import loadUi
import pandas as pd
import re
import sys
import os
import smtplib
from email.message import EmailMessage
import time


class Welcome(QMainWindow):
	def __init__(self):
		super(Welcome, self).__init__()
		loadUi("welcome.ui", self)
		self.b1.clicked.connect(self.gotoInstruction)

	def gotoInstruction(self):
		ins = Instruction()
		widget.addWidget(ins)
		widget.setCurrentIndex(widget.currentIndex()+1)


class Instruction(QMainWindow):
	def __init__(self):
		super(Instruction, self).__init__()
		loadUi("instruction.ui", self)
		self.b22.clicked.connect(self.gotoDemo)
		self.b21.clicked.connect(self.gotoWelcome)
		self.cb.clicked.connect(self.check)

	def check(self):
		if self.cb.isEnabled():
			self.b22.setEnabled(True)

	def gotoWelcome(self):
		widget.setCurrentIndex(widget.currentIndex()-1)

	def gotoDemo(self):
		demo = Demo()
		widget.addWidget(demo)
		widget.setCurrentIndex(widget.currentIndex()+1)


class Demo(QMainWindow):
	def __init__(self):
		super(Demo, self).__init__()
		loadUi("demo.ui", self)
		self.b32.clicked.connect(self.gotoUpload)
		self.b31.clicked.connect(self.gotoInstruction)

	def gotoInstruction(self):
		widget.setCurrentIndex(widget.currentIndex()-1)

	def gotoUpload(self):
		upl = Upload()
		widget.addWidget(upl)
		widget.setCurrentIndex(widget.currentIndex()+1)


class Upload(QMainWindow):
	def __init__(self):
		super(Upload, self).__init__()
		loadUi("upload.ui", self)
		self.b42.clicked.connect(self.gotoCred)
		self.b41.clicked.connect(self.gotoDemo)
		self.b43.clicked.connect(self.browsefilese)
		self.b44.clicked.connect(self.browsefilest)
		self.b45.clicked.connect(self.setstatus)

	def setstatus(self):
		if self.tb1.text() and self.tb2.text() and self.tb3.text():
			global exc
			global col
			global body	
			global rec
			exc = pd.read_excel(excel)
			exc["Year"] = exc["Year"].apply(str)
			col = list(exc.columns)
			body = open(tf, 'r').read()
			rec = []

			try:
				if exc[self.tb3.text()].str.contains("@").all():
					rec = exc[self.tb3.text()]
					self.status.setText("Files uploaded")
					self.b42.setEnabled(True)
				else:
					self.status.setText("Invalid Emails/Emails not entered")
			except:
				self.status.setText("Invalid Email Column!")

		else:
			self.status.setText("Files not uploaded/Email Column Missing")

	def browsefilese(self):
		fname1=QFileDialog.getOpenFileName(self, "Open file", "C:", 'Excel Files (*.xlsx)')
		# print(fname1[0])
		self.tb1.setText(fname1[0])
		global excel
		excel = fname1[0]

	def browsefilest(self):
		fname2=QFileDialog.getOpenFileName(self, "Open file", "C:", 'Text Files (*.txt)')
		# print(fname2[0])
		self.tb2.setText(fname2[0])
		global tf 
		tf = fname2[0]

	def gotoDemo(self):
		widget.setCurrentIndex(widget.currentIndex()-1)

	def gotoCred(self):
		cred = Cred()
		widget.addWidget(cred)
		widget.setCurrentIndex(widget.currentIndex()+1)


class Cred(QMainWindow):
	def __init__(self):
		super(Cred, self).__init__()
		loadUi("cred.ui", self)
		self.b52.clicked.connect(self.gotoMain)
		self.b51.clicked.connect(self.gotoUpload)
		self.b53.clicked.connect(self.setstatus)
		self.tbc2.setEchoMode(QLineEdit.Password)

	def setstatus(self):
		if self.tbc1.text() and self.tbc2.text() and self.tbc3.text():
			string = self.tbc1.text()
			if "@" in string:
				self.status.setText("Credentials entered")
				self.b52.setEnabled(True)
				global email
				global password
				global subject
				email = self.tbc1.text()
				password = self.tbc2.text()
				subject = self.tbc3.text()
			else:
				self.status.setText("Invalid Email")

		else:
			self.status.setText("Credentials not entered")


	def gotoUpload(self):
		widget.setCurrentIndex(widget.currentIndex()-1)

	def gotoMain(self):
		main = Main()
		widget.addWidget(main)
		widget.setCurrentIndex(widget.currentIndex()+1)


class Main(QMainWindow):
	def __init__(self):
		super(Main, self).__init__()
		loadUi("main.ui", self)
		self.b62.clicked.connect(self.gotoPreview)
		self.b61.clicked.connect(self.gotoCred)
		global phf
		global selection
		global p
		pat = r'(?<=\[&#).+?(?=\#&])'
		ph = re.findall(pat, body)
		phf = []
		for i in ph:
			if i not in phf:
				phf.append(i)
		selection = []

		self.cb1.addItems(phf)
		self.cb2.addItems(col)
		self.b63.clicked.connect(self.select)
		self.cb2.activated.connect(self.clicker)
		

	def select(self):
		selection.append(self.cb2.currentText())
		self.cb2.clear()
		self.cb2.addItems(col)
		p = self.cb1.currentIndex()
		if p == len(phf)-1:
			self.show.setText("Selection Done")
			self.cb1.clear()
			self.cb2.clear()
		else:
			self.cb1.setCurrentIndex(p+1)
		self.b63.setEnabled(False)

	def clicker(self):
		self.show.setText(f'You selected: {self.cb1.currentText()}~{self.cb2.currentText()}')
		self.b63.setEnabled(True)
	
	def gotoCred(self):
		widget.setCurrentIndex(widget.currentIndex()-1)

	def gotoPreview(self):
		try:
			preview = Preview()
			widget.addWidget(preview)
			widget.setCurrentIndex(widget.currentIndex()+1)
		except:
			self.show.setText("Please select the placeholders!")


class Preview(QMainWindow):
	def __init__(self):
		super(Preview, self).__init__()
		loadUi("preview.ui", self)
		self.b72.clicked.connect(self.gotoSend)
		self.b71.clicked.connect(self.gotoMain)

		for i in range(1):
			string=body
			for j in range(len(phf)):
				string = re.sub(r'(?<=\[&#){0}?(?=\#&])'.format(phf[j]),exc[selection[j]][i],string)
        
		string = string.replace("[&#", "")
		string = string.replace("#&]", "")	

		self.tb.setText(string)
		self.tb.setFont(QFont('Segoe UI', 12))

	def gotoMain(self):
		widget.setCurrentIndex(widget.currentIndex()-1)

	def gotoSend(self):
		send = Send()
		widget.addWidget(send)
		widget.setCurrentIndex(widget.currentIndex()+1)


class Send(QMainWindow):
	def __init__(self):
		super(Send, self).__init__()
		loadUi("send.ui", self)
		self.b82.clicked.connect(self.gotoStatus)
		self.b81.clicked.connect(self.gotoPreview)
		self.b83.clicked.connect(self.SendMail)

	def SendMail(self):
		try:
			sender = email
			pw = password
			sj = subject

			msg = EmailMessage()
			msg["Subject"] = sj
			msg["From"] = sender

			for i in range(len(selection)):	
				string = body
				for j in range(len(phf)):
					string = re.sub(r'(?<=\[&#){0}?(?=\#&])'.format(phf[j]),exc[selection[j]][i],string)
					if j == len(phf)-1:
						string = string.replace("[&#", "")
						string = string.replace("#&]", "")
						msg["To"] = rec[i]
						msg.set_content(string)
						with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
							smtp.login(sender, pw)
							smtp.send_message(msg)
						print("{0} mail(s) sent!".format(i+1))
						del msg['To']
			self.l.setText("All mails sent!")		
		except:
			self.l.setText("Incorrect Credentials!")
			self.b81.setEnabled(False)



	def gotoPreview(self):
		widget.setCurrentIndex(widget.currentIndex()-1)

	def gotoStatus(self):
		sys.exit(app.exec_())
		# status = Status()
		# widget.addWidget(status)
		# widget.setCurrentIndex(widget.currentIndex()+1)


# class Status(QMainWindow):
# 	def __init__(self):
# 		super(Status, self).__init__()
# 		loadUi("status.ui", self)
# 		self.b92.clicked.connect(self.exit)
# 		self.b91.clicked.connect(self.gotoSend)

# 	def gotoSend(self):
# 		widget.setCurrentIndex(widget.currentIndex()-1)

# 	def exit(self):
# 		sys.exit(app.exec_())


app = QApplication(sys.argv)
widget = QtWidgets.QStackedWidget()
win = Welcome()
widget.addWidget(win)
widget.setFixedHeight(500)
widget.setFixedWidth(800)
widget.show()
widget.setWindowTitle("AutoMail")


try:
	sys.exit(app.exec_())
except:
	print("Exiting")


