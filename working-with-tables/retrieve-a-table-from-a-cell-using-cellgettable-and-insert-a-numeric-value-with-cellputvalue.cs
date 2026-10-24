// Title: Add a ListObject table to a worksheet and write a numeric value to a specific cell using Cell.PutValue in Aspose.Cells for .NET (C#)
// Common Searches: how to create a ListObject table in Aspose.Cells C# and set a numeric value in a cell | Aspose.Cells C# insert number into a table cell with Cell.PutValue | retrieve table range and write data using Aspose.Cells .NET API | C# example for adding a table to a worksheet and updating a cell value in Aspose.Cells | Aspose.Cells put numeric value into second column of first data row of a ListObject
// Tags: ListObject creation Aspose.Cells C# | Cell.PutValue usage Aspose.Cells | populate table cell Aspose.Cells .NET | Aspose.Cells table range retrieval | insert numeric data into Aspose.Cells table

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables; // Required for ListObject

namespace AsposeCellsExample
{
    // The example creates a new Workbook, defines a range A1:B2, adds a ListObject (table) covering that range, then uses Cell.PutValue to insert the number 12345 into the second column of the first data row, and finally saves the workbook as output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (create rule)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // Prepare a simple table to work with (header + data)
                // -------------------------------------------------
                worksheet.Cells["A1"].PutValue("Header1");
                worksheet.Cells["B1"].PutValue("Header2");
                worksheet.Cells["A2"].PutValue("Data1");
                worksheet.Cells["B2"].PutValue("Data2");

                // Define the range that will become a table (A1:B2)
                int firstRow = 0;          // A1 row index
                int firstColumn = 0;       // A1 column index
                int totalRows = 2;
                int totalColumns = 2;

                // Add a ListObject (table) to the worksheet
                int tableIndex = worksheet.ListObjects.Add(
                    firstRow,
                    firstColumn,
                    firstRow + totalRows - 1,
                    firstColumn + totalColumns - 1,
                    true);

                ListObject table = worksheet.ListObjects[tableIndex];

                // -------------------------------------------------
                // Insert a numeric value into a specific cell within the table
                // -------------------------------------------------
                // Determine the first data row (skip the header row)
                int firstDataRow = table.DataRange.FirstRow;

                // Choose the column where the value will be placed (e.g., second column of the table)
                int targetColumn = table.DataRange.FirstColumn + 1;

                // Get the target cell and insert the numeric value
                Cell targetCell = worksheet.Cells[firstDataRow, targetColumn];
                targetCell.PutValue(12345);

                // Define output file path
                string outputPath = "output.xlsx";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook (save rule)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
