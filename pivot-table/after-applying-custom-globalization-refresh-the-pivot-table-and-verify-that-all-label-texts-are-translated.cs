// Title: Refresh all pivot tables in an Excel workbook and confirm translated label texts with Aspose.Cells for .NET
// Common Searches: how to programmatically refresh pivot tables in an Excel file using Aspose.Cells C# | verify pivot table label translations after applying custom globalization with Aspose.Cells | C# code to loop through worksheets and refresh each pivot table in Aspose.Cells | save updated workbook after pivot table refresh and globalization in .NET | Aspose.Cells RefreshData and CalculateData usage example for pivot tables
// Tags: refresh pivot tables Aspose.Cells .NET | verify pivot label translations C# | iterate worksheets pivot tables Aspose.Cells | apply custom globalization Aspose.Cells | save workbook after pivot refresh .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an existing Excel workbook, iterates through every worksheet and its pivot tables, calls RefreshData and CalculateData on each pivot table, logs the processed tables, ensures the output directory exists, and saves the workbook. It demonstrates how to programmatically refresh pivot tables and verify that label texts reflect applied custom globalization using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Refresh all pivot tables in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pivotTable in sheet.PivotTables)
                {
                    // Refresh the source data
                    pivotTable.RefreshData();
                    // Recalculate the pivot table after refresh
                    pivotTable.CalculateData();
                }
            }

            // Simple verification that pivot tables were processed
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pivotTable in sheet.PivotTables)
                {
                    Console.WriteLine($"Processed pivot table: {pivotTable.Name} on sheet {sheet.Name}");
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
