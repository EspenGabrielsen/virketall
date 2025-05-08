import openpyxl
import csv
from tkinter import filedialog as filedialog
from tkinter.messagebox import showwarning
import json


class virketall:
    def __init__(self):
        self.wb = None
        self.ws = None
        self.config = None
        self.choosenFile = None

        self.getConfigFile()
        if self.config:
            self.getFileByname()
        if  self.choosenFile:
            self.readExcel()
            self.startRow = self.config.get("startrow") or None
            self.columns = self.config.get("columns") or None
            self.numberOfCategories = self.config.get("numberOfCategories") or None
            self.categoryNameColumn = self.config.get("categoryNameColumn") or None
            self.headingsOffsetFromStartRow = (
                self.config.get("headingsOffsetFromStartRow") or None
            )
        if self.wb and self.ws:
            self.getValues()
            print(self.valuesList)
    def getConfigFile(self):
        try:
            with open("./config.json", encoding="utf8") as f:
                try:
                    data = json.load(f)
                    self.config = data
                except ValueError:
                    showwarning(
                        title="Parsing error",
                        message="Could not parse the JSON config, it needs to be fixed",
                    )
                return
        except FileNotFoundError:
            showwarning(
                "Could not find config file",
                message="Config file not found please ensure its present",
            )

    def checkIfFileIsChoosen(self):
        if self.choosenFile:
            pass
        else:
            self.getFileByname()

    def getFileByname(self):
        self.choosenFile = filedialog.askopenfilename(
            defaultextension=".xlsx",
            filetypes=(("Excel", "*.xlsx"),),
            title="Choose file with new virketall",
        )

    def readExcel(self):
        self.wb = openpyxl.load_workbook(self.choosenFile) or None
        self.ws = self.wb[self.config.get("sheetName")] or None

    def getValues(self):
        self.valuesList = []
        for tableStartPoint in self.startRow:
            for column in self.columns:
                headers = [
                    self.ws[f"{column}{tableStartPoint+ headerNr}"].value
                    for headerNr in self.headingsOffsetFromStartRow
                ]
                for i in range(self.numberOfCategories):
                    currentRow = tableStartPoint + i
                    currentValue = self.ws[f"{column}{currentRow}"].value
                    currentCategory = self.ws[f"{self.categoryNameColumn}{currentRow}"].value
                    self.addValuesToList(headers=headers,currentValue=currentValue,currentCategory=currentCategory)
    
    def addValuesToList(self,headers,currentValue,currentCategory):
        tempLst = []
        [tempLst.append(header) for header in headers]
        tempLst.append(currentValue)
        tempLst.append(currentCategory)
        self.valuesList.append(tempLst)


if __name__ == "__main__":
    virketall()
