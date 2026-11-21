// Title: Delete a specific row from an Excel ListObject (table) and automatically adjust linked chart series using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# delete a row from a ListObject and keep chart data ranges synchronized | how to remove a table row and update chart series formulas with Aspose.Cells .NET | C# example for deleting the second data row of an Excel table and refreshing chart series using Aspose.Cells | verify chart series ranges after deleting a ListObject row in Aspose.Cells | Aspose.Cells code to delete a table row and automatically adjust linked chart data
// Tags: listobject row deletion aspose.cells | chart series range refresh after table change | c# aspose.cells delete table row | excel chart data update aspose.cells | cells.deleterow listobject example

using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The sample loads a workbook, accesses a ListObject named "Table1", deletes its second data row using Cells.DeleteRow, iterates through all charts on the worksheet to display each series' XValues and Values formulas (confirming they reflect the deletion), and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            Workbook workbook;
            try
            {
                // Load the workbook that contains the ListObject (table) and chart.
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Get the first worksheet (adjust index or name as needed).
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the ListObject (table) by its name. Replace "Table1" with the actual name.
            ListObject listObject = sheet.ListObjects["Table1"];
            if (listObject == null)
            {
                Console.WriteLine("ListObject 'Table1' not found.");
                return;
            }

            // Delete a specific data row from the ListObject (second data row, zero‑based index 1).
            // Use the worksheet's Cells.DeleteRow to remove the row from the underlying range.
            int dataStartRow = listObject.DataRange.FirstRow; // corrected property
            int rowToDelete = dataStartRow + 1; // second data row

            try
            {
                sheet.Cells.DeleteRow(rowToDelete);
                Console.WriteLine($"Deleted row index {rowToDelete} from ListObject 'Table1'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to delete row: {ex.Message}");
                return;
            }

            // Verify that chart series formulas have been updated after the row deletion.
            foreach (Chart chart in sheet.Charts)
            {
                Console.WriteLine($"Chart: {chart.Name}");

                for (int i = 0; i < chart.NSeries.Count; i++)
                {
                    Series series = chart.NSeries[i];
                    // XValues and Values contain the range formulas.
                    string xValues = series.XValues;
                    string values = series.Values;

                    Console.WriteLine($" Series {i}:");
                    Console.WriteLine($"  XValues: {xValues}");
                    Console.WriteLine($"  Values: {values}");
                }
            }

            // Ensure the output directory exists.
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
