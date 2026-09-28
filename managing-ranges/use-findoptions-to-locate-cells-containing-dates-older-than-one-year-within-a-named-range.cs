// Title: Locate Excel cells with dates older than one year inside a named range using Aspose.Cells FindOptions in C#
// AI Prompts: Write C# code that employs Aspose.Cells FindOptions to search a named range for cells whose DateTime values are earlier than a date calculated as one year before today, and return their addresses. | Show how to configure a FindOptions object with a custom date predicate for a named range in Aspose.Cells, then list or highlight the cells that contain dates older than the one‑year cutoff.
// Common Searches: how to use Aspose.Cells FindOptions to filter dates in a specific named range | C# Aspose.Cells find cells with dates before a given cutoff in Excel | search for old dates within a named range using Aspose.Cells .NET | Aspose.Cells date comparison inside named range example | retrieve cell addresses of dates older than one year with Aspose.Cells FindOptions
// Tags: Aspose.Cells FindOptions date filter | C# locate old dates in named range | Excel named range date search Aspose.Cells | Aspose.Cells date cutoff comparison | find cells before specific date .NET

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example loads an Excel workbook, obtains the named range "MyDateRange", iterates through each cell in that range, checks whether the cell holds a DateTime value, compares it to a cutoff date set to one year ago, outputs the address of any cell containing an older date, and finally saves the workbook.
class FindOldDatesInNamedRange
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range
            AsposeRange namedRange = workbook.Worksheets.GetRangeByName("MyDateRange");
            if (namedRange == null)
            {
                Console.WriteLine("Named range 'MyDateRange' not found.");
                return;
            }

            // Determine the bounds of the range
            int firstRow = namedRange.FirstRow;
            int firstColumn = namedRange.FirstColumn;
            int rowCount = namedRange.RowCount;
            int columnCount = namedRange.ColumnCount;
            Worksheet sheet = namedRange.Worksheet;

            // Scan each cell in the range for dates older than one year
            DateTime cutoffDate = DateTime.Now.AddYears(-1);
            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < columnCount; j++)
                {
                    Cell cell = sheet.Cells[firstRow + i, firstColumn + j];
                    if (cell.Type == CellValueType.IsDateTime)
                    {
                        DateTime cellDate = cell.DateTimeValue;
                        if (cellDate < cutoffDate)
                        {
                            Console.WriteLine($"Found old date at: {cell.Name}");
                        }
                    }
                }
            }

            // Save the workbook (optional)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
