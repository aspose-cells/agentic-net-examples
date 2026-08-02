// Title: Find dates older than one year in a named range using FindOptions (Aspose.Cells for .NET)
// Description: Creates a workbook, defines a named range, configures FindOptions with a matching CellArea, scans for DateTime cells earlier than one year, prints their addresses, and saves the file.
// Keywords: Aspose.Cells | FindOptions | named range | date comparison | C# | .NET | CellArea | old dates | Excel automation
// Common Searches: Aspose.Cells FindOptions date older than one year | search named range for dates Aspose.Cells C# | filter DateTime cells in Excel using Aspose | retrieve cells before specific date Aspose.Cells | set search area with CellArea Aspose.Cells
// Developer Intent: Locate and list cells that contain DateTime values older than one year within a specific named range.
// Use Cases: Audit contracts with expiration dates beyond a year | Clean up stale transaction timestamps in financial reports | Generate compliance reports for overdue dates | Flag legacy entries in budgeting models
// AI Prompts: Generate C# code that uses Aspose.Cells FindOptions to directly return addresses of cells with DateTime values earlier than a given cutoff inside a named range. | Show how to replace the manual nested loops with worksheet.Cells.FindAll and a custom predicate for old dates. | Explain the steps to configure FindOptions.SetRange with a CellArea derived from a named range and retrieve matching cells.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Creates a workbook, defines a named range, configures FindOptions with a matching CellArea, scans for DateTime cells earlier than one year, prints their addresses, and saves the file.
class FindOldDates
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data with dates
            worksheet.Cells["A1"].PutValue("Date");
            worksheet.Cells["A2"].PutValue(DateTime.Now.AddMonths(-6));               // 6 months ago
            worksheet.Cells["A3"].PutValue(DateTime.Now.AddYears(-2));                // 2 years ago
            worksheet.Cells["A4"].PutValue(DateTime.Now.AddYears(-1).AddDays(-1));    // just over a year ago
            worksheet.Cells["A5"].PutValue(DateTime.Now);                            // today

            // Create a named range that includes the date cells (A2:A5)
            AsposeRange dateRange = worksheet.Cells.CreateRange("A2:A5");
            dateRange.Name = "DateRange";

            // Retrieve the named range via the workbook
            AsposeRange namedRange = workbook.Worksheets.GetRangeByName("DateRange");

            // Configure FindOptions to limit the search to the named range
            FindOptions findOptions = new FindOptions();
            CellArea searchArea = new CellArea
            {
                StartRow = namedRange.FirstRow,
                StartColumn = namedRange.FirstColumn,
                EndRow = namedRange.FirstRow + namedRange.RowCount - 1,
                EndColumn = namedRange.FirstColumn + namedRange.ColumnCount - 1
            };
            findOptions.SetRange(searchArea);

            // Define the cutoff date (one year ago from today)
            DateTime cutoffDate = DateTime.Now.AddYears(-1);

            // Iterate through cells within the search area and locate dates older than one year
            for (int row = searchArea.StartRow; row <= searchArea.EndRow; row++)
            {
                for (int col = searchArea.StartColumn; col <= searchArea.EndColumn; col++)
                {
                    Cell cell = worksheet.Cells[row, col];
                    if (cell.Type == CellValueType.IsDateTime)
                    {
                        DateTime cellDate = cell.DateTimeValue;
                        if (cellDate < cutoffDate)
                        {
                            Console.WriteLine($"Old date found at {cell.Name}: {cellDate:d}");
                        }
                    }
                }
            }

            // Save the workbook
            string outputPath = "OldDates.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
