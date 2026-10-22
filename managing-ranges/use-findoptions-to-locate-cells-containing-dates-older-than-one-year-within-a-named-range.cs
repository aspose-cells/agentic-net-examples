// Title: Locate Excel cells with dates older than one year inside a named range using Aspose.Cells for .NET
// Common Searches: aspnet find Excel cells older than one year in a named range using Aspose.Cells | c# iterate through a named range and filter dates before today minus 1 year with Aspose.Cells | how to read date values from a named range and compare to a threshold date in Aspose.Cells .NET | search for DateTime cells older than 12 months in an Excel named range with Aspose.Cells | retrieve cells containing past dates from a named range in a workbook using Aspose.Cells C#
// Tags: Aspose.Cells iterate named range cells | Aspose.Cells date threshold comparison | C# filter Excel dates older than one year | named range cell type detection Aspose | Excel workbook date validation using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example loads an Excel workbook, obtains a named range, iterates each cell in that range, checks for DateTime values, compares each date to a one‑year‑old threshold, and outputs the address and value of any dates older than the threshold.
class FindOldDatesInNamedRange
{
    static void Main()
    {
        string inputPath = "input.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the named range
            string rangeName = "MyRange";
            AsposeRange namedRange = workbook.Worksheets.GetRangeByName(rangeName);
            if (namedRange == null)
            {
                Console.WriteLine($"Named range \"{rangeName}\" not found.");
                return;
            }

            int startRow = namedRange.FirstRow;
            int startColumn = namedRange.FirstColumn;
            int rowCount = namedRange.RowCount;
            int columnCount = namedRange.ColumnCount;

            DateTime threshold = DateTime.Now.AddYears(-1);

            // Scan each cell within the named range
            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < columnCount; j++)
                {
                    Cell cell = worksheet.Cells[startRow + i, startColumn + j];
                    if (cell.Type == CellValueType.IsDateTime)
                    {
                        DateTime cellDate = cell.DateTimeValue;
                        if (cellDate < threshold)
                        {
                            Console.WriteLine($"Found old date at {cell.Name}: {cellDate:d}");
                        }
                    }
                }
            }

            // (Optional) Save the workbook if modifications were made
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
