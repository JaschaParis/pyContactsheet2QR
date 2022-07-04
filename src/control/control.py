
import pandas as pd
import numpy as np
import segno.helpers as sh
import segno
import datetime
import os


class Controller:

    def __init__(self):
        self.status_text = "Bitte Daten laden."
        self.contactData = pd.DataFrame(data=None)
        self.dataMapping = {"name": "---", "firstname": "---", "url": "---", "title": "---", "org": "---", "zip": "---",
                        "city": "---", "address": "---", "jobtitle": "---", "photouri": "---", "phone": "---",
                        "cell": "---", "mail": "---"}

    def get_data_labels(self):
        return self.contactData.keys()

    def get_status_text(self):
        return self.status_text

    def load_contact_data(self, filename):
        try:
            self.contactData = pd.read_excel(filename, dtype=str)
        except Exception as e:
            self.status_text = "Datei lesen fehlgeschlagen. \n" + str(e)
            return -1
        contactCount = self.contactData.shape[0]
        self.status_text = "Einträge geladen: " + str(contactCount)
        return 0

    def set_contact_qr_data_mapping(self, data_type_name, mapped_contact_data):
        self.dataMapping[data_type_name] = mapped_contact_data

    def contactrow_has_information(self, row, type):
        if self.dataMapping[type]=="---":
            return False
        elif pd.isna(row[self.dataMapping[type]]):
            return False
        else:
            return True

    def contactrow_get_data(self, row, type):
        if self.contactrow_has_information(row, type):
            return row[self.dataMapping[type]]
        else:
            return None

    def export_qr_codes(self, directory):
        if np.any(self.contactData[self.dataMapping["name"]].isna()):
            self.status_text = "Fehler. Nicht alle Namen gefüllt."
            return -1
        if np.any(self.contactData[self.dataMapping["firstname"]].isna()):
            self.status_text = "Fehler. Nicht alle Vornamen gefüllt."
            return -1

        for index, row in self.contactData.iterrows():
            if not os.path.isdir(directory):
                os.mkdir(directory)
            output_filename = os.path.join(directory, row.Name + "_" + row.Vorname + ".svg")

            name = row[self.dataMapping["name"]]
            firstname = row[self.dataMapping["firstname"]]
            if self.contactrow_has_information(row, "title"):
                completename = name + ";" + row[self.dataMapping["title"]] + " " + firstname
                displayname = row[self.dataMapping["title"]] + " " + firstname + " " + name
            else:
                completename = name + ";" + firstname
                displayname = firstname + " " + name

            phones = []
            if self.contactrow_has_information(row, "phone"):
                phones.append(row[self.dataMapping["phone"]])
            if self.contactrow_has_information(row, "cell"):
                phones.append(row[self.dataMapping["cell"]])

            zipcode = None
            if self.contactrow_has_information(row, "zip"):
                if len(self.contactrow_get_data(row, "zip"))>4:
                    zipcode = self.contactrow_get_data(row, "zip")[0:5]

            city = None
            if self.contactrow_has_information(row, "city"):
                city = self.contactrow_get_data(row, "city")
                if city[0:5].isdecimal():
                    city = city[6:]

            qrstr = sh.make_vcard_data(completename,
                                       displayname,
                                       email=self.contactrow_get_data(row, "mail"),
                                       phone=phones,
                                       url=self.contactrow_get_data(row, "url"),
                                       street=self.contactrow_get_data(row, "address"),
                                       city=city,
                                       zipcode=zipcode,
                                       org=self.contactrow_get_data(row, "org"),
                                       rev=datetime.datetime.now(),
                                       title=self.contactrow_get_data(row, "jobtitle"),
                                       photo_uri=self.contactrow_get_data(row, "photouri"))

            if self.contactrow_has_information(row, "phone") and self.contactrow_has_information(row, "cell"):
                qrstr = qrstr.replace('TEL:', 'TEL;TYPE=WORK:', 1)
                qrstr = qrstr.replace('TEL:', 'TEL;TYPE=CELL:', 1)
            elif self.contactrow_has_information(row, "phone"):
                qrstr = qrstr.replace('TEL:', 'TEL;TYPE=WORK:', 1)
            elif self.contactrow_has_information(row, "cell"):
                qrstr = qrstr.replace('TEL:', 'TEL;TYPE=CELL:', 1)

            qr = segno.make(qrstr.encode())
            qr.save(output_filename, scale=5)
        self.status_text = "Exportieren erfolgreich."
