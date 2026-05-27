using System;
using Aspose.Cells;

namespace AsposeCellsPasswordDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Add some sample data
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Password protected workbook.");

            // ----- Opening password (file encryption) -----
            // This password is required to open the file
            wb.Settings.Password = "OpenPwd123";

            // ----- Modifying (write) password -----
            // This password is required to modify the file after it is opened
            wb.Settings.WriteProtection.Password = "ModifyPwd456";

            // Save the workbook (XLS format)
            string filePath = "PasswordProtectedWorkbook.xls";
            wb.Save(filePath);

            // ------------------- Verification -------------------
            // Load the workbook using the opening password
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = "OpenPwd123";
            Workbook loadedWb = new Workbook(filePath, loadOptions);

            // Check that the write protection password is set
            bool isWriteProtected = loadedWb.Settings.WriteProtection.IsWriteProtected;
            Console.WriteLine("Is write-protected: " + isWriteProtected);

            // Validate the modify password
            bool isValidModifyPwd = loadedWb.Settings.WriteProtection.ValidatePassword("ModifyPwd456");
            Console.WriteLine("Modify password valid: " + isValidModifyPwd);

            // Output a cell value to confirm successful load
            Console.WriteLine("Cell A1 value: " + loadedWb.Worksheets[0].Cells["A1"].Value);
        }
    }
}