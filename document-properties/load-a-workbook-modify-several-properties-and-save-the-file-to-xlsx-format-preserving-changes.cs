// Title: Load, Modify Properties, Style, Settings, Add Worksheet, and Save as XLSX using Aspose.Cells for .NET (C#)
// Description: This C# example shows how to open an existing Excel file with Aspose.Cells, update built‑in properties (Author, Title), add a custom property, change the default font, enable iterative calculation, insert a new worksheet with a timestamp, and save the workbook as an XLSX file while preserving all modifications.
// Keywords: Aspose.Cells C# workbook load | modify Excel document properties Aspose | add custom property Aspose.Cells | default font Aspose.Cells | iterative calculation Aspose.Cells | add worksheet Aspose.Cells | save as XLSX Aspose.Cells | Excel metadata update .NET | Aspose.Cells example
// Common Searches: How to change Author property in Excel using Aspose.Cells C# | Add custom document property to workbook with Aspose.Cells | Set default font for all sheets in Aspose.Cells | Enable iterative calculation in Aspose.Cells | Save modified workbook as XLSX with Aspose.Cells | Aspose.Cells add new worksheet programmatically
// Developer Intent: Programmatically update workbook metadata, style, calculation options, add content, and export to XLSX.
// Use Cases: Prepare compliance‑ready Excel files by updating Author and Title metadata. | Apply corporate branding by setting a default font across the workbook. | Support complex formulas that require iterative calculation before distribution. | Generate summary sheets with timestamps for automated reporting. | Automate bulk updates of Excel files in a .NET workflow.
// AI Prompts: Write C# code using Aspose.Cells to change built‑in properties, add a custom property, set default font, enable iterative calculation, add a worksheet with current date, and save as XLSX. | Provide an Aspose.Cells .NET snippet that updates workbook settings and preserves them after saving. | Explain how to ensure custom document properties are retained when exporting to XLSX with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // This C# example shows how to open an existing Excel file with Aspose.Cells, update built‑in properties (Author, Title), add a custom property, change the default font, enable iterative calculation, insert a new worksheet with a timestamp, and save the workbook as an XLSX file while preserving all modifications.
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file (can be .xlsx, .xls, .csv, etc.)
            string sourcePath = "input.xlsx";

            // Load the workbook using the string constructor (provided rule)
            Workbook workbook = new Workbook(sourcePath);

            // ----- Modify workbook properties -----

            // 1. Built‑in document properties
            workbook.BuiltInDocumentProperties["Author"].Value = "John Doe";
            workbook.BuiltInDocumentProperties["Title"].Value = "Modified Workbook";

            // 2. Custom document property
            workbook.CustomDocumentProperties.Add("ReviewedBy", "Jane Smith");

            // 3. Default style (change default font)
            workbook.DefaultStyle.Font.Name = "Calibri";
            workbook.DefaultStyle.Font.Size = 11;

            // 4. Workbook settings (enable iterative calculation as an example)
            workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;
            workbook.Settings.FormulaSettings.MaxIteration = 50;
            workbook.Settings.FormulaSettings.MaxChange = 0.001;

            // 5. Add a new worksheet and put some data
            int newSheetIndex = workbook.Worksheets.Add();
            Worksheet newSheet = workbook.Worksheets[newSheetIndex];
            newSheet.Name = "Summary";
            newSheet.Cells["A1"].PutValue("Report generated on:");
            newSheet.Cells["B1"].PutValue(DateTime.Now);

            // ----- Save the modified workbook as XLSX -----
            // Use the Save(string, SaveFormat) overload (provided rule)
            workbook.Save("output.xlsx", SaveFormat.Xlsx);

            // Clean up
            workbook.Dispose();

            Console.WriteLine("Workbook loaded, modified, and saved as 'output.xlsx'.");
        }
    }
}
