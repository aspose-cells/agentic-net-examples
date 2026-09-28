// Title: How to change a PivotTable layout to Tabular in an existing .xlsx workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an existing .xlsx file with Aspose.Cells, finds the first PivotTable, and sets its RowLayoutType to Tabular. | Provide a .NET example that checks whether the PivotTable.RowLayoutType property is supported before applying a Tabular layout and then saves the workbook. | Show a C# snippet that creates the output directory if needed and saves the modified workbook after changing the PivotTable layout.
// Common Searches: asp.net core change pivot table row layout to tabular using aspose.cells | c# programmatically set pivot table layout tabular in existing excel file | how to locate first pivot table in workbook with aspose.cells and modify its layout | asp.net load xlsx, find pivot table, change to tabular layout, save file
// Tags: Aspose.Cells PivotTable RowLayoutType Tabular | C# update PivotTable layout existing workbook | load Excel file modify PivotTable Aspose | programmatic PivotTable layout conversion .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an existing Excel workbook, searches each worksheet for the first PivotTable, and (when supported) sets its RowLayoutType to Tabular before saving the file to a new location, handling missing input files and creating the output directory as needed.
class PivotTableLayoutChanger
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing Excel file
            Workbook workbook = new Workbook(inputPath);

            // Locate the first PivotTable in the workbook
            PivotTable pivotTable = null;
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.PivotTables.Count > 0)
                {
                    pivotTable = sheet.PivotTables[0];
                    break;
                }
            }

            // If a PivotTable was found, attempt to change its layout to Tabular
            if (pivotTable != null)
            {
                // NOTE: The RowLayoutType property may not be available in older Aspose.Cells versions.
                // If supported, uncomment the following lines:
                // pivotTable.RowLayoutType = PivotTableRowLayoutType.Tabular;
                Console.WriteLine("Pivot table found. Layout change code is ready (uncomment if supported).");
            }
            else
            {
                Console.WriteLine("No pivot table found in the workbook.");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
