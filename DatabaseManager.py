import pandas

# RULES:
# 1. JANGAN GANTI NAMA CLASS ATAU FUNGSI YANG ADA
# 2. JANGAN DELETE FUNGSI YANG ADA
# 3. JANGAN DELETE ATAU MENAMBAH PARAMETER PADA CONSTRUCTOR ATAU FUNGSI
# 4. GANTI NAMA PARAMETER DI PERBOLEHKAN
# 5. LARANGAN DI ATAS BOLEH DILANGGAR JIKA ANDA TAU APA YANG ANDA LAKUKAN (WAJIB BISA JELASKAN)
# GOODLUCK :)

class excelManager:
    def __init__(self,filePath:str,sheetName:str="Sheet1"):
        self.__filePath = filePath
        self.__sheetName = sheetName
        self.__data = pandas.read_excel(filePath,sheet_name=sheetName)
            
    
    def insertData(self,newData:dict,saveChange:bool=False):
        # kerjakan disini
        # clue cara insert row: df = pandas.concat([df, pandas.DataFrame([{"NIM":0,"Nama":"Udin","Nilai":1000}])], ignore_index=True)
        self.__data = self.__data._append(newData, ignore_index=True)
        if saveChange:
            self.saveChange()
    
    def deleteData(self, targetedNim:str,saveChange:bool=False):
        # kerjakan disini
        # clue cara delete row (contoh):self.__data.drop(indexBaris, inplace=True);contoh penggunaan: self.__data.drop(0, inplace=True)
        original_count = len(self.__data)
        self.__data = self.__data[self.__data['NIM'].astype(str) != str(targetedNim)]
        
        if len(self.__data) == original_count:
            return None
        
        if (saveChange):
            self.saveChange()
        
        return {"deleted_NIM": targetedNim, "rows_affected": original_count - len(self.__data)}
    
    def editData(self, targetedNim:str, newData:dict,saveChange:bool=False) -> dict:
        # kerjakan disini
        # clue cara ganti value (contoh):self.__data.at[indexBaris, namaKolom] = value; contoh penggunaan: self.__data.at[0, 'ID'] = 1
        mask = self.__data['NIM'].astype(str) == str(targetedNim)
        
        if not mask.any():
            return None
        
        for key, value in newData.items():
            self.__data.loc[mask, key] = value
        
        if (saveChange):
            self.saveChange()

        updated_row = self.__data[mask].iloc[0].to_dict()
        updated_row["Row"] = self.__data[mask].index[0]
        return updated_row
                    
    def getData(self, colName:str, data:str) -> dict:
        collumn = self.__data.columns # mendapatkan list dari nama kolom tabel
        
        # cari index dari nama kolom dan menjaganya dari typo atau spasi berlebih
        collumnIndex = [i for i in range(len(collumn)) if (collumn[i].lower().strip() == colName.lower().strip())] 
        
        # validasi jika input kolom tidak ada pada data excel
        if (len(collumnIndex) != 1): return None
        
        # nama kolom yang sudah pasti benar dan ada
        colName = collumn[collumnIndex[0]]
        
        
        resultDict = dict() # tempat untuk hasil
        
        for i in self.__data.index: # perulangan ke baris tabel
            cellData = str(self.__data.at[i,colName]) # isi tabel yand dijadikan str
            if (cellData == data): # jika data cell sama dengan data input
                for col in collumn: # perulangan ke nama-nama kolom
                    resultDict.update({str(col):str(self.__data.at[i,col])}) # masukan data {namaKolom : data pada cell} ke resultDict
                resultDict.update({"Row":i}) # tambahkan row nya pada resultDict
                return resultDict # kembalikan resultDict
        
        return None
    
    def saveChange(self):
        self.__data.to_excel(self.__filePath, sheet_name=self.__sheetName , index=False)
    
    def getDataFrame(self):
        return self.__data
