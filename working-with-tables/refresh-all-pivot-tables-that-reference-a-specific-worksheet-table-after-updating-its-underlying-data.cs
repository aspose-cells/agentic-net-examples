// Title: Refresh every PivotTable linked to a specific ListObject after modifying its data with Aspose.Cells for .NET (C#)
// AI Prompts: Modify a cell in a named ListObject and then invoke RefreshData and CalculateData on all PivotTables in the workbook using Aspose.Cells C#. | Write C# code that loops through each worksheet, refreshes each PivotTable’s cache, and saves the updated workbook to a new file. | Add robust error‑handling for missing worksheets, missing tables, and pivot refresh failures when working with Aspose.Cells.
// Common Searches: aspnet refresh all pivot tables after updating an Excel table with Aspose.Cells | c# programmatically refresh pivot cache when ListObject data changes | how to recalculate pivot tables in a workbook after editing a table using Aspose.Cells | iterate worksheets to refresh pivot tables Aspose.Cells .NET example | save workbook after pivot table refresh Aspose.Cells C#
// Tags: Aspose.Cells refresh pivot tables C# | update ListObject data Aspose.Cells | pivot cache refresh after table edit | iterate worksheets Aspose.Cells | error handling missing worksheet Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Pivot;

// The example loads a workbook, updates a cell in a named ListObject, iterates through all worksheets to refresh each PivotTable’s cache and recalculate its data, and then saves the modified workbook to a new file while handling possible missing worksheet or table errors.
class RefreshPivotTables
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Locate the worksheet that contains the table (adjust the name as needed)
            Worksheet dataSheet = workbook.Worksheets["Data"];
            if (dataSheet == null)
            {
                Console.WriteLine("Worksheet 'Data' not found.");
                return;
            }

            // Retrieve the table (ListObject) by its name (adjust the name as needed)
            ListObject table = dataSheet.ListObjects["MyTable"];
            if (table == null)
            {
                Console.WriteLine("Table 'MyTable' not found.");
                return;
            }

            // ----- Update underlying data of the table -----
            // Example: modify the first data row, first column (excluding header)
            int firstDataRow = table.DataRange.FirstRow;
            int firstDataColumn = table.DataRange.FirstColumn;
            dataSheet.Cells[firstDataRow, firstDataColumn].PutValue(12345); // new value

            // ----- Refresh all pivot tables in the workbook -----
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pivot in sheet.PivotTables)
                {
                    try
                    {
                        // Refresh the pivot cache data and recalculate the pivot table
                        pivot.RefreshData();
                        pivot.CalculateData();
                    }
                    catch (Exception exPivot)
                    {
                        Console.WriteLine($"Failed to refresh pivot table '{pivot.Name}' on sheet '{sheet.Name}': {exPivot.Message}");
                    }
                }
            }

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the updated workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception exSave)
            {
                Console.WriteLine($"Failed to save workbook: {exSave.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
