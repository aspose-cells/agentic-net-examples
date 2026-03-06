using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Set the author for write protection
        workbook.Settings.WriteProtection.Author = "John Doe";

        // (Optional) Set a password and recommend read‑only mode
        workbook.Settings.WriteProtection.Password = "ownerPassword";
        workbook.Settings.WriteProtection.RecommendReadOnly = true;

        // Save the workbook with the write‑protection settings applied
        string outputPath = "output_protected.xlsx";
        workbook.Save(outputPath);
    }
}