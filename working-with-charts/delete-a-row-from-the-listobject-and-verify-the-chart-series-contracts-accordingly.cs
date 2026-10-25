// Title: Delete a row from an Excel ListObject and ensure chart series ranges are updated with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells delete a table row and keep chart series references updated | How to remove a data row from a ListObject and automatically adjust chart ranges in .NET | Verify chart series formulas after deleting a row from an Excel table using Aspose.Cells | Aspose.Cells delete rows from worksheet cells and refresh chart NSeries | Update chart data source after removing a row from an Excel table with Aspose.Cells
// Tags: Aspose.Cells ListObject row deletion | automatic chart series range adjustment | C# worksheet.Cells.DeleteRows chart update | Excel table modification affecting NSeries | chart data source verification Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Charts;

// The example loads an Excel workbook, checks for a ListObject, deletes a specific data row from the first table, iterates through all charts on the worksheet to display each series' name and its XValues/YValues formulas, and saves the workbook, demonstrating that chart series ranges are automatically updated after the row removal.
class DeleteRowAndVerifyChart
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

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Verify that a ListObject (table) exists
            if (worksheet.ListObjects.Count == 0)
            {
                Console.WriteLine("No ListObject found in the worksheet.");
                return;
            }

            // Get the first ListObject
            ListObject listObject = worksheet.ListObjects[0];

            // Determine which data row to delete (e.g., the second data row)
            // ListObject.DataRange includes the header row, so add 1 to skip header
            int rowToDelete = listObject.DataRange.FirstRow + 1; // zero‑based index

            // Delete the row from the worksheet
            worksheet.Cells.DeleteRows(rowToDelete, 1);

            // Verify that chart series ranges have been updated accordingly
            Console.WriteLine("Verifying chart series after row deletion:");
            foreach (Chart chart in worksheet.Charts)
            {
                Console.WriteLine($"Chart: {chart.Name ?? "(unnamed)"}");
                foreach (Series series in chart.NSeries)
                {
                    // Series name (if defined)
                    string seriesName = !string.IsNullOrEmpty(series.Name) ? series.Name : "(no name)";

                    // XValues range formula
                    string xValuesFormula = !string.IsNullOrEmpty(series.XValues) ? series.XValues : "(no X values)";

                    // YValues range formula
                    string yValuesFormula = !string.IsNullOrEmpty(series.Values) ? series.Values : "(no Y values)";

                    Console.WriteLine($"  Series:");
                    Console.WriteLine($"    Name      : {seriesName}");
                    Console.WriteLine($"    XValues   : {xValuesFormula}");
                    Console.WriteLine($"    YValues   : {yValuesFormula}");
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
