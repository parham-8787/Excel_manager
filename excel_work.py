from openpyxl import Workbook , load_workbook
class Excel :
    def _save(self):
        """Save workbook with correct extension"""
        self.wb.save(fr"{self.addres_file}.xlsx" if not self.addres_file.endswith('.xlsx') else fr"{self.addres_file}")
    

    def __init__(self,addres_file,sheet_name,*,create_workbook="no"):
        if create_workbook == "yes":
            try:
                self.addres_file=addres_file
                self.wb=load_workbook(fr"{addres_file}.xlsx" if not addres_file.endswith('.xlsx') else fr"{addres_file}")
                self.ws=self.wb[sheet_name]
            except FileNotFoundError:
                print("File not found, creating a new workbook...")
                self.wb = Workbook()
                self.ws = self.wb.active
            except Exception as e :
                print(e)
            else:
                self._save()
                print("done")

        elif create_workbook == "no":
            try:
                self.addres_file=addres_file
                self.wb = Workbook()
                self.ws=self.wb.active
            except Exception as e:
                print(e)
            else:
                self._save()
                print("Workbook Created...")
        else:
            print("Enter a valid value")


    def create_sheet(self,sheet_name):
        try:
            self.wb.create_sheet(title=sheet_name)
        except Exception as e:
            print(e)
        else:
            self._save()
            print("Sheet Created...")


    def rename_sheet(self,sheet_name,new_sheet_name):
            try:
                sheet = self.wb[sheet_name]
                sheet.title = new_sheet_name
            except Exception as e:
                print(e)
            else:
                self._save()
                print("Sheet renamed...")


    def remove_sheet(self,sheet_name):
        try:
            sheet = self.wb[sheet_name]
            self.wb.remove(sheet)
        except Exception as e:
            print(e)
        else:
            self._save()
            print("Sheet removed...")


    def show_sheets(self):
        try:
            sheets_name=self.wb.sheetnames
        except Exception as e:
            print(e)
        else:
            for n in sheets_name:
                print(n,end=" | ")



    def insert_data(self,data):
        try:
            if isinstance(data[0],(list,tuple)):
                for i in range(len(data)):
                    self.ws.append(data[i])
            else:
                self.ws.append(data)
        except Exception as e:
            print(e)
        else:
            self._save()
            print("data appended...")
    
    def show_data(self,min_row,max_row,min_column,max_column):
        try:
            data_show=[]
            for row in self.ws.iter_rows(min_row=min_row,max_row=max_row,min_col=min_column,max_col=max_column,values_only=True):
                data_show.append(row)
                print(*row)
        except Exception as e:
            print(e)
        else:
            return data_show
        
    def excel_to_excel(self,min_row,max_row,min_col,max_col,target_file):
        try:
            wb2 =load_workbook(fr"{target_file}.xlsx" if not target_file.endswith('.xlsx') else fr"{target_file}")
            ws2=wb2.active
            for row in self.ws.iter_rows(min_row=min_row,max_row=max_row,min_col=min_col,max_col=max_col,values_only=True):
                ws2.append(row)
            wb2.save(fr"{target_file}.xlsx" if not target_file.endswith('.xlsx') else fr"{target_file}")
        except FileNotFoundError:
                try:
                    print("File not found, creating a new workbook...")
                    wb2 = Workbook()
                    ws2 = wb2.active
                    for row in self.ws.iter_rows(min_row=min_row,max_row=max_row,min_col=min_col,max_col=max_col,values_only=True):
                        ws2.append(row)
                    wb2.save(fr"{target_file}.xlsx" if not target_file.endswith('.xlsx') else fr"{target_file}")
                except Exception as e:
                    print(e)
        except Exception as e:
            print(e)
        else:
            print("ok")
            