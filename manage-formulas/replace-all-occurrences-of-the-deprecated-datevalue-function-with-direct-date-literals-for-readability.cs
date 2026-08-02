// Title: Replace DATEVALUE with literal dates in Aspose.Cells for .NET (C#)
// Description: C# example that creates a workbook, inserts DATEVALUE formulas, scans all cells, evaluates each DATEVALUE expression, substitutes the formula with the resulting Excel date serial number, applies a standard date format, and saves the file. Ideal for modernising legacy spreadsheets using Aspose.Cells.
// Keywords: Aspose.Cells C# | Aspose.Cells .NET | replace DATEVALUE | date serial number | convert DATEVALUE to literal date | calculate formula Aspose.Cells | Excel date format | legacy spreadsheet migration | GitHub Aspose.Cells example | Excel compatibility
// Common Searches: how to remove DATEVALUE function with Aspose.Cells | Aspose.Cells replace DATEVALUE with date value | C# convert DATEVALUE formula to serial number | Aspose.Cells calculate and replace formulas | update old Excel sheets that use DATEVALUE
// Developer Intent: Detect DATEVALUE formulas, evaluate them, write the resulting serial number as a plain value, and format the cell as a date.
// Use Cases: Modernise workbooks that contain deprecated DATEVALUE functions for better compatibility with newer Excel versions. | Improve calculation performance by storing dates as static serial numbers instead of formulas. | Prepare spreadsheets for export to systems that do not support Excel functions such as DATEVALUE.
// AI Prompts: Show C# code using Aspose.Cells to find and replace DATEVALUE formulas with their evaluated serial numbers. | Explain how to handle multiple date string formats when converting DATEVALUE formulas in Aspose.Cells. | Provide a step‑by‑step guide to apply a date number format after replacing DATEVALUE in a workbook.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // C# example that creates a workbook, inserts DATEVALUE formulas, scans all cells, evaluates each DATEVALUE expression, substitutes the formula with the resulting Excel date serial number, applies a standard date format, and saves the file. Ideal for modernising legacy spreadsheets using Aspose.Cells.
    public class ReplaceDateValueFunctionDemo
    {
        public static void Main(string[] args)
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
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data: formulas that use the deprecated DATEVALUE function
            // =DATEVALUE("2021-01-01")  -> should become the serial number for 2021-01-01
            // =DATEVALUE("12/31/2022")  -> should become the serial number for 31-Dec-2022
            cells["A1"].Formula = "=DATEVALUE(\"2021-01-01\")";
            cells["A2"].Formula = "=DATEVALUE(\"12/31/2022\")";

            // Iterate through all used cells in the worksheet
            foreach (Cell cell in cells)
            {
                // Process only cells that contain a formula with DATEVALUE
                if (cell.IsFormula && cell.Formula.IndexOf("DATEVALUE", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // Calculate the result of the existing DATEVALUE formula
                    object result = sheet.CalculateFormula(cell.Formula);

                    // Replace the formula with the calculated result (Excel date serial number)
                    cell.PutValue(result);

                    // Apply a built‑in date format for readability
                    Style style = cell.GetStyle();
                    style.Number = 14; // mm/dd/yyyy
                    cell.SetStyle(style);
                }
            }

            // Define output file path
            string outputPath = "ReplaceDateValueDemo.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
