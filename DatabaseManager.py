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
        if not isinstance(newData, dict):
            raise TypeError("newData must be a dict")
        new_row_df = pandas.DataFrame([newData])
        self.__data = pandas.concat([self.__data, new_row_df], ignore_index=True)
        if (saveChange):
            self.saveChange()
    
    def deleteData(self, targetedNim:str,saveChange:bool=False):
        # kerjakan disini
        # clue cara delete row (contoh):self.__data.drop(indexBaris, inplace=True);contoh penggunaan: self.__data.drop(0, inplace=True)
        if 'NIM' not in self.__data.columns:
            return None
        mask = self.__data['NIM'].astype(str) == str(targetedNim)
        matches = self.__data[mask]
        if matches.empty:
            return None
        deleted_rows = []
        for _, row in matches.iterrows():
            row_dict = {str(col): str(row[col]) for col in self.__data.columns}
            deleted_rows.append(row_dict)
        self.__data.drop(matches.index, inplace=True)

        if (saveChange):
            self.saveChange()
        return deleted_rows[0] if len(deleted_rows) == 1 else deleted_rows
    
    def editData(self, targetedNim:str, newData:dict,saveChange:bool=False) -> dict:
        # kerjakan disini
        # clue cara ganti value (contoh):self.__data.at[indexBaris, namaKolom] = value; contoh penggunaan: self.__data.at[0, 'ID'] = 1
        if not isinstance(newData, dict):
            raise TypeError("newData must be a dict")
        if 'NIM' not in self.__data.columns:
            return None
        mask = self.__data['NIM'].astype(str) == str(targetedNim)
        matches = self.__data[mask]
        if matches.empty:
            return None
        for idx in matches.index:
            for key, val in newData.items():
                self.__data.at[idx, key] = val

        if (saveChange):
            self.saveChange()

        first_idx = matches.index[0]
        updated_row = {str(col): str(self.__data.at[first_idx, col]) for col in self.__data.columns}
        updated_row.update({"Row": int(first_idx)})
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
