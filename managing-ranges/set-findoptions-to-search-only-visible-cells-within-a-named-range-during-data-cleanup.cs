// Title: Find visible cells in a named range using Aspose.Cells FindOptions (C#)
// Description: C# example that creates a workbook, defines a named range (A1:B3), hides a row, builds a CellArea from the range, configures FindOptions to search only within that area, and returns the first visible cell matching a text value. The workbook is then saved.
// Keywords: Aspose.Cells FindOptions | visible cells | named range | C# .NET | exclude hidden rows | CellArea search | Find method | data cleanup | Aspose.Cells API | US developers
// Common Searches: Aspose.Cells find only visible cells in a named range | SetRange FindOptions hide hidden rows C# | Search non‑hidden cells with Aspose.Cells Find | How to limit Find to visible cells in Aspose.Cells .NET | Find text in named range while ignoring hidden rows
// Developer Intent: Locate a cell that contains a specific value, but only among visible rows and columns inside a defined named range.
// Use Cases: Validate user input by detecting visible duplicate entries within a scoped range. | Generate a filtered report that lists only visible matches from a large dataset. | Automate data‑cleanup scripts that skip hidden rows/columns while searching for keywords.
// AI Prompts: Write C# code that uses Aspose.Cells FindOptions to search for a string in a named range while ignoring hidden rows and columns. | Explain how to build a CellArea from a named range and apply FindOptions.SetRange to restrict the search to visible cells. | Show how to extend the example to retrieve all matches with FindAll while keeping the visibility filter.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// C# example that creates a workbook, defines a named range (A1:B3), hides a row, builds a CellArea from the range, configures FindOptions to search only within that area, and returns the first visible cell matching a text value. The workbook is then saved.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data
            worksheet.Cells["A1"].PutValue("Apple");
            worksheet.Cells["A2"].PutValue("Banana");
            worksheet.Cells["A3"].PutValue("Cherry");
            worksheet.Cells["B1"].PutValue("Date");
            worksheet.Cells["B2"].PutValue("Elderberry");
            worksheet.Cells["B3"].PutValue("Fig");

            // Hide the second row to simulate invisible cells
            worksheet.Cells.Rows[1].IsHidden = true;

            // Create a named range that covers the whole area
            AsposeRange namedRange = worksheet.Cells.CreateRange("A1", "B3");
            namedRange.Name = "FruitRange";

            // Retrieve the named range via the workbook's worksheet collection
            AsposeRange range = workbook.Worksheets.GetRangeByName("FruitRange");

            // Convert the Range to a CellArea for FindOptions
            CellArea searchArea = new CellArea
            {
                StartRow = range.FirstRow,
                StartColumn = range.FirstColumn,
                EndRow = range.FirstRow + range.RowCount - 1,
                EndColumn = range.FirstColumn + range.ColumnCount - 1
            };

            // Configure FindOptions to search only within the defined range
            FindOptions findOptions = new FindOptions
            {
                LookInType = LookInType.Values,
                LookAtType = LookAtType.Contains,
                SearchOrderByRows = true
            };
            findOptions.SetRange(searchArea);

            // Perform a find operation for the word "Apple"
            Cell foundCell = worksheet.Cells.Find("Apple", null, findOptions);

            // Ensure the found cell is not hidden (visible only)
            if (foundCell != null &&
                !worksheet.Cells.Rows[foundCell.Row].IsHidden &&
                !worksheet.Cells.Columns[foundCell.Column].IsHidden)
            {
                Console.WriteLine($"Found visible cell: {foundCell.Name} with value '{foundCell.StringValue}'");
            }
            else
            {
                Console.WriteLine("No visible cell found within the named range.");
            }

            // Save the workbook
            workbook.Save("FindVisibleInNamedRange.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
