using System;
using Aspose.Cells;

namespace AsposeCellsWriteProtectionRemoval
{
    class Program
    {
        static void Main()
        {
            // Path to the write‑protected workbook
            string inputPath = "WriteProtectedWorkbook.xlsx";

            // Load the workbook. If it was protected with a password,
            // provide the password via LoadOptions so the file can be opened.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = "owner"; // replace with actual password if needed
            Workbook wb = new Workbook(inputPath, loadOptions);

            // Remove write‑protection by clearing the password and related settings
            wb.Settings.WriteProtection.Password = null;          // no password
            wb.Settings.WriteProtection.Author = null;           // optional: clear author
            wb.Settings.WriteProtection.RecommendReadOnly = false; // optional: clear read‑only flag

            // Save the workbook without any write‑protection
            string outputPath = "UnprotectedWorkbook.xlsx";
            wb.Save(outputPath);

            // Verify removal (optional)
            Workbook verifyWb = new Workbook(outputPath);
            Console.WriteLine("Is workbook write protected? " +
                verifyWb.Settings.WriteProtection.IsWriteProtected);
        }
    }
}