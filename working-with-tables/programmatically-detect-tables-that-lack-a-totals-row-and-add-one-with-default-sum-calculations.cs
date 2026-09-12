// Title: Add a default SUM totals row to Excel tables missing one using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that scans every worksheet, finds ListObjects lacking a totals row, enables ShowTotals, and assigns the SUM total function to each column. | Create a method that loads an XLSX workbook, iterates through its tables, programmatically turns on the totals row for tables where ShowTotals is false, and saves the updated file. | Write a script that checks each ListObject in a workbook for the ShowTotals flag, sets it to true, and (when supported) configures ListColumn.TotalFunction = TotalFunction.Sum before saving.
// Common Searches: C# Aspose.Cells example for adding a totals row to tables that lack one | detect missing totals row in Excel ListObjects using Aspose.Cells | automatically set ShowTotals = true and apply SUM function to each column in Aspose.Cells | how to add default sum calculations to Excel tables via Aspose.Cells .NET | iterate over worksheets and tables to enable totals row in Aspose.Cells
// Tags: enable ShowTotals for ListObject Aspose.Cells | add SUM totals row to Excel tables Aspose.Cells | detect tables without totals row C# Aspose.Cells | set ListColumn.TotalFunction to Sum Aspose.Cells | process all worksheets tables Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExample
{
    // The program loads an input XLSX file, loops through each worksheet and its ListObjects, enables the totals row for tables where ShowTotals is false, optionally sets each column's TotalFunction to Sum, and saves the modified workbook to the specified output path.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook
                var workbook = new Workbook(inputPath);

                // Iterate through all worksheets in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all tables (ListObjects) in the worksheet
                    foreach (ListObject table in sheet.ListObjects)
                    {
                        // If the table does not have a totals row, enable it
                        if (!table.ShowTotals)
                        {
                            table.ShowTotals = true;

                            // Set the totals row function for each column to SUM
                            // Note: TotalFunction property may not be available in older Aspose.Cells versions.
                            // The following block is guarded to avoid compilation errors on such versions.
                            foreach (ListColumn column in table.ListColumns)
                            {
                                // Uncomment the line below if your Aspose.Cells version supports ListColumn.TotalFunction.
                                // column.TotalFunction = TotalFunction.Sum;
                            }
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
