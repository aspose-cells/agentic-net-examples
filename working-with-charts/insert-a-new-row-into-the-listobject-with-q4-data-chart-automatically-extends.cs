// Title: Add a Q4 row to an Excel ListObject and automatically expand the linked chart with Aspose.Cells for .NET
// Common Searches: how to programmatically insert a new row into an Excel ListObject and refresh the chart using Aspose.Cells C# | Aspose.Cells C# add Q4 sales data to an existing table and extend the chart range | resize an Excel table after adding a row so the linked chart updates automatically with Aspose.Cells | C# insert quarterly data into a ListObject and keep the chart data source synchronized using Aspose.Cells | using Aspose.Cells to add a row to a table and automatically adjust chart data source in .NET
// Tags: programmatic ListObject row insertion Aspose.Cells | table resize with header preservation Aspose.Cells | auto‑extend chart data source after table update | add quarterly sales row to Excel workbook Aspose.Cells | maintain chart linkage when expanding Excel table .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, finds the first worksheet and its first ListObject, inserts a new row after the current data range, writes "Q4" and a sales value, resizes the table to include the new row while preserving the header setting, and saves the file. Any chart linked to the table automatically expands to include the new Q4 data.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            try
            {
                // Load the existing workbook that contains the ListObject (table) and the chart.
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index or name as needed).
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the first ListObject on the worksheet.
                // If the table has a specific name, use worksheet.ListObjects["TableName"] instead.
                if (worksheet.ListObjects.Count == 0)
                {
                    Console.WriteLine("Error: No tables (ListObjects) found in the worksheet.");
                    return;
                }

                ListObject table = worksheet.ListObjects[0];

                // Determine the row index where the new data row should be inserted (after the current data range).
                int insertRowIndex = table.DataRange.FirstRow + table.DataRange.RowCount;

                // Insert a new empty row into the worksheet at the calculated position.
                worksheet.Cells.InsertRows(insertRowIndex, 1);

                // Populate the new row with Q4 data.
                // Adjust column indices to match the table's structure.
                worksheet.Cells[insertRowIndex, 0].PutValue("Q4");          // Quarter
                worksheet.Cells[insertRowIndex, 1].PutValue(12345);       // Sales (replace with actual value)

                // Resize the ListObject to include the newly added row.
                bool hasHeaders = table.ShowHeaderRow; // true if the table has a header row
                int firstRow = table.DataRange.FirstRow - (hasHeaders ? 1 : 0);
                int firstColumn = table.DataRange.FirstColumn;
                int totalRows = table.DataRange.RowCount + (hasHeaders ? 1 : 0) + 1; // +1 for the new row
                int totalColumns = table.DataRange.ColumnCount;

                // Resize overload requires the hasHeaders flag.
                table.Resize(firstRow, firstColumn, totalRows, totalColumns, hasHeaders);

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
