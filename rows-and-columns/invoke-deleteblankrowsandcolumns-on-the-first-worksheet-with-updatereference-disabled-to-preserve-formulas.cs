// Title: Delete Blank Rows & Columns on First Worksheet (Preserve Formulas) – Aspose.Cells for .NET
// Description: Loads or creates a workbook, adds sample values and a formula, configures DeleteOptions with UpdateReference = false, removes empty rows and columns from the first worksheet using DeleteBlankRows and DeleteBlankColumns, and saves the file while keeping all formulas intact.
// Keywords: Aspose.Cells delete blank rows | Aspose.Cells delete blank columns | preserve formulas Aspose.Cells | DeleteOptions UpdateReference false | C# Aspose.Cells clean worksheet | remove empty rows .NET | remove empty columns .NET | Aspose.Cells DeleteBlankRows example | Aspose.Cells DeleteBlankColumns example
// Common Searches: how to delete empty rows in Aspose.Cells without breaking formulas | Aspose.Cells DeleteOptions UpdateReference usage | remove blank columns from first worksheet C# | Aspose.Cells keep formulas when deleting rows | DeleteBlankRows DeleteBlankColumns sample code
// Developer Intent: Remove all blank rows and columns from the first worksheet while ensuring existing formulas remain unchanged.
// Use Cases: Sanitize imported spreadsheets by stripping out unused rows/columns before data analysis. | Prepare report templates that retain only populated cells, preserving summary calculations. | Automate a data‑processing pipeline that cleans workbooks without affecting dependent formulas.
// AI Prompts: Show a C# example that deletes blank rows and columns on the first worksheet using Aspose.Cells while keeping formulas by disabling UpdateReference. | Explain how DeleteOptions.UpdateReference influences formula references when calling DeleteBlankRows and DeleteBlankColumns. | Generate code to load a workbook, clean empty rows/columns from sheet 0, and save the result without altering any formulas.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads or creates a workbook, adds sample values and a formula, configures DeleteOptions with UpdateReference = false, removes empty rows and columns from the first worksheet using DeleteBlankRows and DeleteBlankColumns, and saves the file while keeping all formulas intact.
    public class DeleteBlankRowsAndColumnsDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook (or load an existing one if the file exists)
            Workbook workbook;
            string inputPath = "input.xlsx";

            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // start with a blank workbook
            }

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data with blank rows/columns and formulas
            worksheet.Cells["A1"].PutValue(10);
            worksheet.Cells["B1"].PutValue(20);
            worksheet.Cells["C1"].Formula = "=A1+B1"; // Formula that will be preserved
            worksheet.Cells["A3"].PutValue(30); // Row 2 is blank
            worksheet.Cells["B5"].PutValue(50); // Column C is blank

            // Set delete options with UpdateReference disabled (preserve formulas)
            DeleteOptions deleteOptions = new DeleteOptions
            {
                UpdateReference = false
            };

            // Delete blank rows and columns using the options
            worksheet.Cells.DeleteBlankRows(deleteOptions);
            worksheet.Cells.DeleteBlankColumns(deleteOptions);

            // Save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
