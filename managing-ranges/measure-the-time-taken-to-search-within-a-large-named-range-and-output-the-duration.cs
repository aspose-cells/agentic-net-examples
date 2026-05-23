using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

class SearchNamedRangeTiming
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate a large range of data (e.g., 10,000 rows × 10 columns)
            int totalRows = 10000;
            int totalColumns = 10;
            for (int row = 0; row < totalRows; row++)
            {
                for (int col = 0; col < totalColumns; col++)
                {
                    cells[row, col].PutValue(row * col);
                }
            }

            // Create a named range that covers the populated area
            int nameIdx = workbook.Worksheets.Names.Add("LargeRange");
            Name largeRangeName = workbook.Worksheets.Names[nameIdx];

            // Build the address string for the bottom‑right cell (e.g., J10000)
            string endAddress = cells[totalRows - 1, totalColumns - 1].Name;
            largeRangeName.RefersTo = $"=Sheet1!A1:{endAddress}";

            // Retrieve the Range object for the named range (use fully qualified name to avoid ambiguity)
            Aspose.Cells.Range largeRange = largeRangeName.GetRange();

            // Configure FindOptions to limit the search to the named range
            FindOptions findOptions = new FindOptions
            {
                LookInType = LookInType.Values,
                LookAtType = LookAtType.EntireContent
            };

            // Define the CellArea that corresponds to the named range
            CellArea searchArea = new CellArea
            {
                StartRow = largeRange.FirstRow,
                StartColumn = largeRange.FirstColumn,
                EndRow = largeRange.FirstRow + largeRange.RowCount - 1,
                EndColumn = largeRange.FirstColumn + largeRange.ColumnCount - 1
            };
            findOptions.SetRange(searchArea);

            // Value to search for (value at bottom‑right cell)
            object targetValue = (totalRows - 1) * (totalColumns - 1);

            // Measure the time taken to perform the search
            Stopwatch stopwatch = Stopwatch.StartNew();
            Cell foundCell = worksheet.Cells.Find(targetValue, null, findOptions);
            stopwatch.Stop();

            // Output the duration and result
            Console.WriteLine($"Search completed in {stopwatch.ElapsedMilliseconds} ms.");
            if (foundCell != null)
            {
                Console.WriteLine($"Value found at cell {foundCell.Name}.");
            }
            else
            {
                Console.WriteLine("Value not found.");
            }

            // Save the workbook (optional)
            string outputPath = "SearchTiming.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}