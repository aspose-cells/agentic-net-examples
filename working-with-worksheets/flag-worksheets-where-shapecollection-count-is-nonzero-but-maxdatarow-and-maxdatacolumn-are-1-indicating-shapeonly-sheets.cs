// Title: Detect and Flag Shape‑Only Worksheets in Excel with Aspose.Cells for .NET
// Description: Loads an Excel workbook, scans each worksheet, and flags those that contain shapes but no cell data (MaxDataRow = -1 and MaxDataColumn = -1). The flag is added as a custom property "ShapeOnly" and the sheet name is suffixed with "_ShapeOnly" before saving the file.
// Keywords: Aspose.Cells shape only worksheet detection | C# identify worksheets with only drawings | add custom property to Excel sheet Aspose | MaxDataRow MaxDataColumn -1 check | rename worksheet based on shapes Aspose.Cells | flag shape‑only sheets .NET
// Common Searches: how to find Excel sheets that only have shapes using Aspose.Cells | add custom property to mark shape‑only worksheets in .NET | rename worksheets with drawings but no data Aspose.Cells | detect empty data rows and columns in Excel via Aspose
// Developer Intent: Locate worksheets that contain drawings but no cell data and mark them for downstream processing.
// Use Cases: Exclude shape‑only sheets when converting a workbook to PDF. | Automatically rename drawing‑only tabs for easier navigation in large workbooks. | Store a flag in custom properties to drive conditional logic in automation pipelines.
// AI Prompts: Create C# code using Aspose.Cells that scans all worksheets, flags those with Shapes.Count > 0 and MaxDataRow/MaxDataColumn = -1 by adding a "ShapeOnly" custom property and appending "_ShapeOnly" to the sheet name. | Provide an Aspose.Cells snippet that logs the names of shape‑only worksheets and saves the modified workbook. | Write a method that returns a list of worksheet names that have shapes but no data rows or columns using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, scans each worksheet, and flags those that contain shapes but no cell data (MaxDataRow = -1 and MaxDataColumn = -1). The flag is added as a custom property "ShapeOnly" and the sheet name is suffixed with "_ShapeOnly" before saving the file.
class ShapeOnlySheetFlagger
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine if the sheet contains shapes but no data rows/columns
                bool hasShapes = sheet.Shapes.Count > 0;
                bool noDataRows = sheet.Cells.MaxDataRow == -1;
                bool noDataCols = sheet.Cells.MaxDataColumn == -1;

                if (hasShapes && noDataRows && noDataCols)
                {
                    // Flag the worksheet by adding a custom property (value must be a string)
                    sheet.CustomProperties.Add("ShapeOnly", "true");

                    // Optionally, rename the sheet to make the flag visible in Excel
                    sheet.Name = sheet.Name + "_ShapeOnly";
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
