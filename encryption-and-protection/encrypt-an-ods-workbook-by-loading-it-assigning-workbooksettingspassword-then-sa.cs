using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (instead of loading a non‑existent ODS file)
            Workbook workbook = new Workbook();

            // Add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello, Aspose.Cells!");
            sheet.Name = "Demo";

            // Assign a password to encrypt the workbook
            workbook.Settings.Password = "StrongPassword123";

            // Save the encrypted workbook as XLSX format
            string xlsxPath = "encrypted_output.xlsx";
            workbook.Save(xlsxPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved as '{xlsxPath}' with password protection.");
        }
    }
}