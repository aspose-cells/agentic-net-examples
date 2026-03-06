using System;
using Aspose.Cells;

namespace AsposeCellsPasswordModifyDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Add sample data
            Worksheet ws = wb.Worksheets[0];
            ws.Cells["A1"].PutValue("Protected content");

            // Set the write‑protection password (modify option)
            wb.Settings.WriteProtection.Password = "modifyPwd";

            // Save the workbook as XLSX
            string filePath = "WriteProtectedWorkbook.xlsx";
            wb.Save(filePath, SaveFormat.Xlsx);

            // Load the workbook with the password to allow modifications
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = "modifyPwd";
            Workbook wbLoaded = new Workbook(filePath, loadOptions);

            // Perform modifications
            wbLoaded.Worksheets[0].Cells["B2"].PutValue("Edited after password");

            // Remove write‑protection if you want to save without password
            wbLoaded.Settings.WriteProtection.Password = null;

            // Save the modified workbook
            wbLoaded.Save("WriteProtectedWorkbook_Modified.xlsx", SaveFormat.Xlsx);
        }
    }
}