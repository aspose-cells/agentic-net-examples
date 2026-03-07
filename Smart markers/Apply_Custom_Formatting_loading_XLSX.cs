using System;
using Aspose.Cells;

namespace CustomFormattingExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Create LoadOptions and configure desired loading behavior
            LoadOptions loadOptions = new LoadOptions();
            // Example: skip formula parsing to speed up loading (optional)
            loadOptions.ParsingFormulaOnOpen = false;
            // Example: enable auto‑filtering when loading
            loadOptions.AutoFilter = true;

            // Load the workbook with the specified LoadOptions
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Enable strict checking of custom number formats (optional)
            workbook.Settings.CheckCustomNumberFormat = true;

            // Create a custom style with a custom number format
            Style customStyle = workbook.CreateStyle();
            // Custom format: show numbers with thousand separators and two decimal places
            customStyle.Custom = "#,##0.00";

            // Apply the custom style to a specific cell
            Worksheet sheet = workbook.Worksheets[0];
            Cell targetCell = sheet.Cells["B2"];
            targetCell.PutValue(12345.678);   // Sample numeric value
            targetCell.SetStyle(customStyle); // Apply the custom formatting

            // Save the modified workbook to a new file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
        }
    }
}