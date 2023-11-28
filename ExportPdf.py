import win32com.client
import os
import shutil
import time
import subprocess


def get_pdf():
    source_file = r'Location of the File'
    # Define the destination directory (current working directory)
    destination_dir = os.getcwd()

    # Copy the Excel file to the destination directory
    shutil.copy(source_file, destination_dir)

    source_file = os.path.join(destination_dir, "Hardik Patel.xlsm")

    # Create a new Excel application object
    # If you get error in this line, you need to delete the gen_py cache (C:\Users\hpatel\AppData\Local\Temp\gen_py)
    # Error: AttributeError: module 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9' has no attribute 'CLSIDToClassMap'
    excel = win32com.client.gencache.EnsureDispatch("Excel.Application")


    # Open the Excel file that you want to add the module to
    workbook = excel.Workbooks.Open(source_file)
    try:
        time.sleep(2)
        module = workbook.VBProject.VBComponents.Add(1)
        module.CodeModule.AddFromString("Sub ChangeWeekNoFilter()\n" +
                                    "Dim ws As Worksheet\n" +
                                    "Dim filterCell As Range\n" +
                                    "Dim pdfFileName As String\n" +
                                    "Dim exportRange As Range\n\n" +
                                    "' Set the worksheet where you want to change the filter\n" +
                                    "Set ws = ThisWorkbook.Worksheets(\"Time Sheet\")\n" +
                                    "' Set the cell with the ""WEEK NO"" filter\n" +
                                    "Set filterCell = ws.Range(\"B4\")\n" +
                                    "' Set the new filter value\n" +
                                    "currentWeek = ws.Range(\"D4\").Value\n" +
                                    "filterCell.Value = currentWeek - 1\n" +
                                    "' Optionally, trigger an event to refresh the filter\n" +
                                    "Application.SendKeys \"{ENTER}\"\n" +
                                    "' Set the PDF file name\n" +
                                    "pdfFileName = ThisWorkbook.Path & \"\Hardik Patel.pdf\"\n" +
                                    "' Set the custom range (A1 to K45) to export as PDF\n" +
                                    "Set exportRange = ws.Range(\"A1:K45\")\n" +
                                    "' Export the specified range to a PDF file\n" +
                                    "exportRange.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfFileName\n" +
                                    "End Sub")

        # Save the Excel file
        excel.Run("ChangeWeekNoFilter")

        workbook.Save()
        workbook.Close()
        excel.Quit()
        pdf_path = os.path.join(destination_dir, "Hardik Patel.pdf")
        if os.path.isfile(pdf_path):
            try:
                subprocess.Popen(['start', '', pdf_path], shell=True)
            except Exception as e:
                print(f"Error opening the PDF: {e}")

    except Exception as e:
        print(e)
        excel.Quit()

if __name__ == "__main__":
    get_pdf()