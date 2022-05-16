from PyQt5 import QtCore, QtGui, QtWidgets
from google_auth_oauthlib import flow
import datetime as dt
import pandas as pd
import os
import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


class Ui_Dialog(object):
    def setupUi(self, Dialog):
        SCOPES = ["https://www.googleapis.com/auth/calendar"]
        creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'client_secrets.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.json', 'w') as token:
                token.write(creds.to_json())
        self.service = build('calendar', 'v3', credentials=creds)

        Dialog.setObjectName("Dialog")
        Dialog.resize(449, 437)
        self.SubmitButton = QtWidgets.QPushButton(Dialog)
        self.SubmitButton.setGeometry(QtCore.QRect(140, 340, 161, 71))
        self.SubmitButton.setObjectName("SubmitButton")
        self.SubmitButton.clicked.connect(self.submit_form)
        self.SubmitButton.setStyleSheet('QPushButton {background-color: #326b7d; color: white;}')

        self.label = QtWidgets.QLabel(Dialog)
        self.label.setGeometry(QtCore.QRect(120, 30, 231, 41))
        self.label.setStyleSheet("")
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(Dialog)
        self.label_2.setGeometry(QtCore.QRect(130, 90, 60, 16))
        self.label_2.setObjectName("label_2")
        self.label_3 = QtWidgets.QLabel(Dialog)
        self.label_3.setGeometry(QtCore.QRect(130, 140, 60, 16))
        self.label_3.setObjectName("label_3")
        self.label_4 = QtWidgets.QLabel(Dialog)
        self.label_4.setGeometry(QtCore.QRect(90, 190, 111, 16))
        self.label_4.setObjectName("label_4")
        self.label_5 = QtWidgets.QLabel(Dialog)
        self.label_5.setGeometry(QtCore.QRect(90, 240, 111, 16))
        self.label_5.setObjectName("label_5")
        self.label_6 = QtWidgets.QLabel(Dialog)
        self.label_6.setGeometry(QtCore.QRect(90, 290, 111, 16))
        self.label_6.setObjectName("label_6")

        self.DateSelect = QtWidgets.QComboBox(Dialog)
        self.DateSelect.setGeometry(QtCore.QRect(210, 90, 131, 28))
        self.DateSelect.setObjectName("DateSelect")
        self.init_dates()

        self.TimeSelect = QtWidgets.QComboBox(Dialog)
        self.TimeSelect.setGeometry(QtCore.QRect(210, 140, 131, 28))
        self.TimeSelect.setObjectName("TimeSelect")

        self.lineEdit = QtWidgets.QLineEdit(Dialog)
        self.lineEdit.setGeometry(QtCore.QRect(210, 190, 131, 21))
        self.lineEdit.setObjectName("lineEdit")
        self.lineEdit_2 = QtWidgets.QLineEdit(Dialog)
        self.lineEdit_2.setGeometry(QtCore.QRect(210, 240, 131, 21))
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.lineEdit_3 = QtWidgets.QLineEdit(Dialog)
        self.lineEdit_3.setGeometry(QtCore.QRect(210, 290, 131, 21))
        self.lineEdit_3.setObjectName("lineEdit_3")
        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Appointment Scheduler"))
        self.SubmitButton.setText(_translate("Dialog", "SUBMIT"))
        self.label.setText(_translate("Dialog", "Enter customer reservation details:"))
        self.label_2.setText(_translate("Dialog", "Date"))
        self.label_3.setText(_translate("Dialog", "Time"))
        self.label_4.setText(_translate("Dialog", "Customer Name"))
        self.label_5.setText(_translate("Dialog", "Customer Phone"))
        self.label_6.setText(_translate("Dialog", "Customer Email"))
        self.DateSelect.setCurrentText(_translate("Dialog", "Date"))
        self.TimeSelect.setCurrentText(_translate("Dialog", "Time"))

    def init_dates(self):
        self.DateSelect.clear()
        today = dt.date.today()
        four_weeks = today + dt.timedelta(weeks=4)
        date_range = pd.date_range(start=today, end=four_weeks)
        date_list = [x.strftime("%d %B, %Y") for x in date_range]
        self.DateSelect.addItems(date_list)
        self.DateSelect.setCurrentIndex(-1)
        self.DateSelect.currentTextChanged.connect(self.change_time_avail)
        return

    def submit_form(self):
        service = self.service
        date = dt.datetime.strptime(self.DateSelect.currentText(), "%d %B, %Y")
        time = self.TimeSelect.currentText()
        hour = int(time.split(':')[0])
        minute = int(time.split(':')[1])
        start = date + dt.timedelta(hours=hour + 7, minutes=minute)
        end = start + dt.timedelta(minutes=30) - dt.timedelta(seconds=1)
        start = start.isoformat() + 'Z'
        end = end.isoformat() + 'Z'
        name = self.lineEdit.text()
        phone = self.lineEdit_2.text()
        email = self.lineEdit_3.text()
        event = {
            'summary': name + ' - Device Repair Appointment',
            'location': 'Midtown Device Repair LLC.',
            'description': 'An appointment for your device repair has been scheduled! Customer Contact Info: ' + name + ' ' + email + ' ' + phone,
            'start': {
                'dateTime': start,
                'timeZone': 'America/New_York',
            },
            'end': {
                'dateTime': end,
                'timeZone': 'America/New_York',
            },
            'attendees': [
                {'email': 'midtowndevicerepair@gmail.com'},
            ],
            'reminders': {
                'useDefault': False,
                'overrides': [
                    {'method': 'email', 'minutes': 10},
                    {'method': 'popup', 'minutes': 10},
                ],
            },
        }

        event = service.events().insert(calendarId='primary', body=event).execute()
        self.init_dates()
        return

    def change_time_avail(self):
        service = self.service
        try:
            selected = dt.datetime.strptime(self.DateSelect.currentText(), "%d %B, %Y")
        except:
            self.TimeSelect.clear()
            self.lineEdit.clear()
            self.lineEdit_2.clear()
            self.lineEdit_3.clear()
            return
        date_selected = dt.date(selected.year, selected.month, selected.day)
        if date_selected == dt.date.today():
            start_time = dt.datetime.now()
        else:
            start_time = dt.datetime(date_selected.year, date_selected.month,
                                     date_selected.day, hour=8, minute=0)
        end_time = dt.datetime(date_selected.year, date_selected.month,
                               date_selected.day, hour=18, minute=0)

        slots = pd.date_range(dt.datetime(date_selected.year, date_selected.month,
                                          date_selected.day, hour=8, minute=0), end_time,
                              freq='30T')
        slots = slots[[x >= start_time for x in slots]]
        availability = []
        for idx in range(0, len(slots)-1):
            start = slots[idx]
            end = slots[idx+1] - dt.timedelta(seconds=1)
            events_result = service.events().list(calendarId='primary', timeMin=start.isoformat() + 'Z', timeMax=end.isoformat() + 'Z', singleEvents=True,
                                                  orderBy='startTime').execute()
            events = events_result.get('items', [])
            if not events:
                availability.append(start)

        avails = [x.strftime('%H:%M') for x in availability]
        self.TimeSelect.clear()
        self.TimeSelect.addItems(avails)
        return


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Dialog = QtWidgets.QDialog()
    ui = Ui_Dialog()
    ui.setupUi(Dialog)
    Dialog.show()
    sys.exit(app.exec_())
