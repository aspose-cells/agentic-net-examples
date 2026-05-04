using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Paths for the demonstration files
        string sourcePath = "Template.xlsm";          // Existing macro‑enabled workbook
        string copyPath = "CopyWithMacro.xlsm";       // Copy that keeps macros
        string noMacroPath = "CopyNoMacro.xlsx";      // Copy with macros removed
        string finalPath = "FinalWorkbook.xlsx";      // Workbook after data manipulation

        // Ensure the source workbook exists; if not, create a placeholder macro‑enabled file
        if (!File.Exists(sourcePath))
        {
            Workbook placeholder = new Workbook();
            placeholder.Save(sourcePath, SaveFormat.Xlsm);
        }

        // ------------------------------------------------------------
        // 1. Load an existing XLSM workbook and inspect macro presence
        // ------------------------------------------------------------
        Workbook sourceWorkbook = new Workbook(sourcePath);
        Console.WriteLine($"Source workbook has macro: {sourceWorkbook.HasMacro}");

        // ------------------------------------------------------------
        // 2. Copy the workbook while preserving macros
        // ------------------------------------------------------------
        if (sourceWorkbook.HasMacro)
        {
            // Create an empty destination workbook
            Workbook copyWorkbook = new Workbook();

            // Configure copy options to keep macros
            CopyOptions copyOptions = new CopyOptions
            {
                KeepMacros = true
            };

            // Perform the copy operation
            sourceWorkbook.Copy(copyWorkbook, copyOptions);

            // Save the copied workbook as macro‑enabled file
            copyWorkbook.Save(copyPath, SaveFormat.Xlsm);
            Console.WriteLine($"Saved copy with macros to: {copyPath}");
        }

        // ------------------------------------------------------------
        // 3. Remove macros from a workbook and save as macro‑free file
        // ------------------------------------------------------------
        Workbook noMacroWorkbook = new Workbook(sourcePath); // Load again
        noMacroWorkbook.RemoveMacro();                     // Remove all VBA/macros
        noMacroWorkbook.Save(noMacroPath, SaveFormat.Xlsx);
        Console.WriteLine($"Saved macro‑free workbook to: {noMacroPath}");

        // ------------------------------------------------------------
        // 4. Complex data manipulation on a new workbook
        // ------------------------------------------------------------
        Workbook dataWorkbook = new Workbook();               // Create a new workbook
        Worksheet sheet = dataWorkbook.Worksheets[0];
        sheet.Name = "Data";

        // Add header row
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Score");

        // Populate sample data (10 rows)
        for (int row = 2; row <= 11; row++)
        {
            sheet.Cells[$"A{row}"].PutValue(row - 1);                     // ID = 1..10
            sheet.Cells[$"B{row}"].PutValue(100 - (row - 2) * 5);        // Decreasing scores
        }

        // Insert a formula to calculate the average score
        sheet.Cells["B12"].Formula = "AVERAGE(B2:B11)";
        dataWorkbook.CalculateFormula(); // Evaluate the formula

        // Replace a specific cell value (demonstrates Replace/PutValue)
        sheet.Cells["B5"].PutValue(999);

        // Save the manipulated workbook
        dataWorkbook.Save(finalPath, SaveFormat.Xlsx);
        Console.WriteLine($"Saved final workbook with data manipulation to: {finalPath}");
    }
}