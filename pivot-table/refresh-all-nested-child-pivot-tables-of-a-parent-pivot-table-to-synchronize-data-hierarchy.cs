// Title: Refresh all nested child pivot tables in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an XLSX file with Aspose.Cells, iterates through every worksheet, finds each PivotTable (including child tables), calls RefreshData on them, logs any failures, and saves the workbook. | Create a reusable method in C# that accepts a Workbook object and refreshes all descendant pivot tables of a specified parent pivot table, handling missing files and exception logging with Aspose.Cells.
// Common Searches: aspocells .net refresh nested pivot tables in workbook | how to programmatically refresh all child pivot tables using Aspose.Cells C# | iterate through worksheets and refresh each pivot table Aspose.Cells example | C# code to refresh pivot tables and log errors with Aspose.Cells | refresh pivot table hierarchy in Excel file using Aspose.Cells for .NET
// Tags: refresh nested pivot tables Aspose.Cells .NET | iterate worksheets refresh pivot data C# | handle missing Excel file Aspose.Cells | log pivot table refresh errors .NET | bulk pivot table refresh Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExample
{
    // The example loads an existing XLSX workbook with Aspose.Cells, checks that the input file exists, then loops through every worksheet and each PivotTable—including any child tables—calling RefreshData() to synchronize the data hierarchy. It captures and logs any refresh failures, and finally saves the updated workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "Input.xlsx";
            string outputPath = "Output.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                var workbook = new Workbook(inputPath);

                // Refresh all pivot tables in each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (PivotTable pt in sheet.PivotTables)
                    {
                        try
                        {
                            // Refresh the pivot table data
                            pt.RefreshData();
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Failed to refresh pivot table '{pt.Name}' on sheet '{sheet.Name}': {ex.Message}");
                        }
                    }
                }

                // Save the updated workbook
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
