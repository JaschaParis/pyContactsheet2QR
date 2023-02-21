import wx
from control.control import Controller
from view.logo import fulllogo

BORDER_VALUE = 10


class MainFrame(wx.Frame):

    def __init__(self):

        self._frame = wx.Frame.__init__(self, None, wx.ID_ANY,
                          "pyContactsheet2QR", size=(400,900))
        wx.GetApp().SetAppName("pyContactsheet2QR")
        panel = wx.Panel(self)

        mainSizer = wx.BoxSizer(wx.VERTICAL)

        statusSection = wx.StaticBoxSizer(wx.VERTICAL, panel, "Status:")
        loadFileSection = wx.StaticBoxSizer(wx.VERTICAL, panel, "Datei importieren:")
        dataSelectionSection = wx.StaticBoxSizer(wx.VERTICAL, panel, "Zuordnung der Spaltennamen:")
        exportQRcodesSection = wx.StaticBoxSizer(wx.VERTICAL, panel, "Exportiere QR Codes:")

        self.btn_loadFile = wx.Button(panel, id=wx.ID_ANY, label="Datei wählen")
        self.btn_loadFile.Bind(wx.EVT_BUTTON, self.loadData)
        self.btn_exportQRcodes = wx.Button(panel, id=wx.ID_ANY, label="Exportiere QR Codes")
        self.btn_exportQRcodes.Bind(wx.EVT_BUTTON, self.exportQRcodes)

        self.statusText = wx.StaticText(panel, id=wx.ID_ANY, label="")

        statusSection.Add(self.statusText, wx.EXPAND, border=BORDER_VALUE)
        loadFileSection.Add(self.btn_loadFile, wx.EXPAND, wx.ALIGN_CENTER, border=BORDER_VALUE)
        exportQRcodesSection.Add(self.btn_exportQRcodes, wx.EXPAND, wx.ALIGN_CENTER, border=BORDER_VALUE)

        dataSizer = wx.FlexGridSizer(2,13,10)
        dataSizer.AddGrowableCol(1,proportion=4)

        lbl_typ = wx.StaticText(panel, id=wx.ID_ANY, label="Typ:")
        lbl_spalte = wx.StaticText(panel, id=wx.ID_ANY, label="Spaltennamen:")

        lbl_name = wx.StaticText(panel, id=wx.ID_ANY, label="Name:")
        lbl_firstname = wx.StaticText(panel, id=wx.ID_ANY, label="Vorname:")
        lbl_title = wx.StaticText(panel, id=wx.ID_ANY, label="Akad. Titel:")
        lbl_org = wx.StaticText(panel, id=wx.ID_ANY, label="Organisation:")
        lbl_zip = wx.StaticText(panel, id=wx.ID_ANY, label="PLZ:")
        lbl_city = wx.StaticText(panel, id=wx.ID_ANY, label="Stadt:")
        lbl_address = wx.StaticText(panel, id=wx.ID_ANY, label="Adresse:")
        lbl_jobtitle = wx.StaticText(panel, id=wx.ID_ANY, label="Jobtitel:")
        lbl_phone = wx.StaticText(panel, id=wx.ID_ANY, label="Telefon:")
        lbl_cell = wx.StaticText(panel, id=wx.ID_ANY, label="Mobil:")
        lbl_mail = wx.StaticText(panel, id=wx.ID_ANY, label="Mail:")
        lbl_url = wx.StaticText(panel, id=wx.ID_ANY, label="Website:")
        lbl_photouri = wx.StaticText(panel, id=wx.ID_ANY, label="Photo-URI:")

        self.cb_name = wx.Choice(panel, id=wx.ID_ANY, name="name")
        self.cb_firstname = wx.Choice(panel, id=wx.ID_ANY, name="firstname")
        self.cb_title = wx.Choice(panel, id=wx.ID_ANY, name="title")
        self.cb_org = wx.Choice(panel, id=wx.ID_ANY, name="org")
        self.cb_zip = wx.Choice(panel, id=wx.ID_ANY, name="zip")
        self.cb_city = wx.Choice(panel, id=wx.ID_ANY, name="city")
        self.cb_address = wx.Choice(panel, id=wx.ID_ANY, name="address")
        self.cb_jobtitle = wx.Choice(panel, id=wx.ID_ANY, name="jobtitle")
        self.cb_phone = wx.Choice(panel, id=wx.ID_ANY, name="phone")
        self.cb_cell = wx.Choice(panel, id=wx.ID_ANY, name="cell")
        self.cb_mail = wx.Choice(panel, id=wx.ID_ANY, name="mail")
        self.cb_url = wx.Choice(panel, id=wx.ID_ANY, name="url")
        self.cb_photouri = wx.Choice(panel, id=wx.ID_ANY, name="photouri")

        self.nameMatching = {"name": ["Name", "Nachname", "Surname"],
                             "firstname": ["Vorname", "Firstname", "First name"],
                             "url": ["Website", "Url"], "title": ["Titel", "Akad"],
                             "org": ["Gesellschaft", "Organisation", "Firma"],
                             "zip": ["PLZ", "Postleitzahl", "zip"],
                             "city": ["Stadt", "Ort", "City"],
                             "address": ["Straße"],
                             "jobtitle": ["Job", "Position", "Funktion"],
                             "photouri": ["Photo", "Foto"],
                             "phone": ["Festnetz", "Durchwahl", "Telefon"],
                             "cell": ["Handy", "Mobil", "Cell"],
                             "mail": ["E-Mail", "Mail"]}

        self.list_cb = [self.cb_name, self.cb_firstname, self.cb_url, self.cb_title, self.cb_org, self.cb_zip,
                        self.cb_city, self.cb_address, self.cb_jobtitle, self.cb_photouri, self.cb_phone,
                        self.cb_cell, self.cb_mail]

        dataSizer.Add(lbl_typ, 0, wx.EXPAND, wx.ALIGN_CENTER)
        dataSizer.Add(lbl_spalte, 0, wx.EXPAND, wx.ALIGN_CENTER)
        dataSizer.Add(lbl_name, 0, wx.EXPAND)
        dataSizer.Add(self.cb_name, 2, wx.EXPAND)
        dataSizer.Add(lbl_firstname, 0, wx.EXPAND)
        dataSizer.Add(self.cb_firstname, 2, wx.EXPAND)
        dataSizer.Add(lbl_title, 0, wx.EXPAND)
        dataSizer.Add(self.cb_title, 2, wx.EXPAND)
        dataSizer.Add(lbl_org, 0, wx.EXPAND)
        dataSizer.Add(self.cb_org, 2, wx.EXPAND)
        dataSizer.Add(lbl_jobtitle, 0, wx.EXPAND)
        dataSizer.Add(self.cb_jobtitle, 2, wx.EXPAND)
        dataSizer.Add(lbl_address, 0, wx.EXPAND)
        dataSizer.Add(self.cb_address, 2, wx.EXPAND)
        dataSizer.Add(lbl_zip, 0, wx.EXPAND)
        dataSizer.Add(self.cb_zip, 2, wx.EXPAND)
        dataSizer.Add(lbl_city, 0, wx.EXPAND)
        dataSizer.Add(self.cb_city, 2, wx.EXPAND)
        dataSizer.Add(lbl_phone, 0, wx.EXPAND)
        dataSizer.Add(self.cb_phone, 2, wx.EXPAND)
        dataSizer.Add(lbl_cell, 0, wx.EXPAND)
        dataSizer.Add(self.cb_cell, 2, wx.EXPAND)
        dataSizer.Add(lbl_mail, 0, wx.EXPAND)
        dataSizer.Add(self.cb_mail, 2, wx.EXPAND)
        dataSizer.Add(lbl_url, 0, wx.EXPAND)
        dataSizer.Add(self.cb_url, 2, wx.EXPAND)
        dataSizer.Add(lbl_photouri, 0, wx.EXPAND)
        dataSizer.Add(self.cb_photouri, 2, wx.EXPAND)

        dataSelectionSection.Add(dataSizer, 1, wx.EXPAND, border=BORDER_VALUE)

        img = fulllogo.GetBitmap()
        bitmap = wx.StaticBitmap(panel, id=wx.ID_ANY, bitmap=img, size=(img.GetWidth(), img.GetHeight()))
        mainSizer.Add(bitmap, 0)

        mainSizer.Add(statusSection, 1, wx.EXPAND, border=BORDER_VALUE)
        mainSizer.Add(loadFileSection, 1, wx.EXPAND, border=BORDER_VALUE)
        mainSizer.Add(dataSelectionSection, 8, wx.EXPAND, border=BORDER_VALUE)
        mainSizer.Add(exportQRcodesSection, 1, wx.EXPAND, border=BORDER_VALUE)

        panel.SetSizer(mainSizer)

        self.Layout()
        self.Show()
        self.Centre()

        self.SetSize((400,220))
        dataSelectionSection.ShowItems(False)
        exportQRcodesSection.ShowItems(False)
        self.dataSelectionSection = dataSelectionSection
        self.exportQRcodesSection = exportQRcodesSection

        self.ctrl = Controller()
        self.statusText.SetLabelText(self.ctrl.get_status_text())

    def loadData(self, evt):
        chooseFile = wx.FileDialog(self._frame, "Wähle Datei",
                                         wildcard="xlsx Dateien (*.xlsx)|*.xlsx",
                                         style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
        if chooseFile.ShowModal() == wx.ID_CANCEL:
            return
        fileChoosen = chooseFile.GetPath()

        if self.ctrl.load_contact_data(fileChoosen) == 0:
            self.updateComboBoxText()
            self.dataSelectionSection.ShowItems(True)
            self.exportQRcodesSection.ShowItems(True)
            self.SetSize((400,900))
            self.Layout()
            self.Show()
            self.Centre()
        self.statusText.SetLabelText(self.ctrl.get_status_text())

    def updateComboBoxText(self):
        dataLabels = self.ctrl.get_data_labels()

        for cb in self.list_cb:
            cb.SetItems(["---"])
            cb.SetStringSelection("---")
            for dataLabel in dataLabels:
                cb.AppendItems(dataLabel)
                if any(match in dataLabel for match in self.nameMatching[cb.GetName()]):
                    cb.SetStringSelection(dataLabel)

    def updateControllerContactData(self):
        for cb in self.list_cb:
            self.ctrl.set_contact_qr_data_mapping(cb.GetName(), cb.GetStringSelection())

    def exportQRcodes(self, evt):
        if self.cb_name.GetStringSelection()=="---":
            wxmodal = wx.MessageDialog(self, "Es muss ein Datenfeld für den Namen ausgewählt sein.", "Warnung", style=wx.OK | wx.ICON_WARNING)
            wxmodal.ShowModal()
            return

        if self.cb_firstname.GetStringSelection()=="---":
            wxmodal = wx.MessageDialog(self, "Es muss ein Datenfeld für den Vornamen ausgewählt sein.", "Warnung", style=wx.OK | wx.ICON_WARNING)
            wxmodal.ShowModal()
            return

        chooseDir = wx.DirDialog(self._frame, 'Wähle Exportordner',
                                         style=wx.DD_DIR_MUST_EXIST)  # correct style chosen?
        if chooseDir.ShowModal() == wx.ID_CANCEL:
            return
        choosenDirectory = chooseDir.GetPath()
        if choosenDirectory:
            self.updateControllerContactData()
            self.ctrl.export_qr_codes(choosenDirectory)
            self.statusText.SetLabelText(self.ctrl.get_status_text())
