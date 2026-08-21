// Title: Find visible cells in a named range using FindOptions – Aspose.Cells for .NET
// Description: Creates a workbook, hides a row and a column, defines a named range (A1:B5), builds a CellArea, configures FindOptions to limit the search to that area, iterates with Cells.Find to locate the value "Active" only in cells that are not hidden, and saves the result to an XLSX file.
// Keywords: Aspose.Cells FindOptions | visible cells search | named range lookup | exclude hidden rows | exclude hidden columns | .NET spreadsheet API | CellArea range
// Common Searches: Aspose.Cells find only visible cells | search named range ignoring hidden rows | FindOptions SetRange example .NET | how to skip hidden columns in Aspose.Cells search | retrieve named range by name Aspose.Cells
// Developer Intent: Locate a specific value within a named range while ignoring any hidden rows or columns.
// Use Cases: Validate data entries that are displayed to the user, skipping hidden rows. | Generate a list of active items from a filtered view of a worksheet. | Implement a cleanup routine that processes only visible cells in a defined range.
// AI Prompts: Show how to configure FindOptions to search a named range and exclude hidden rows and columns in Aspose.Cells for .NET. | Provide a short code snippet that uses FindOptions.SetRange with a CellArea and filters out hidden cells when searching for a value. | Explain the steps to retrieve a named range by name and use it with Cells.Find to locate visible cells only.

using System;
using System.IO;
using Aspose.Cells;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

// Creates a workbook, hides a row and a column, defines a named range (A1:B5), builds a CellArea, configures FindOptions to limit the search to that area, iterates with Cells.Find to locate the value "Active" only in cells that are not hidden, and saves the result to an XLSX file.
class FindVisibleInNamedRange
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            cells["A1"].PutValue("Item");
            cells["B1"].PutValue("Status");
            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue("Active");
            cells["A3"].PutValue("Banana");
            cells["B3"].PutValue("Inactive");
            cells["A4"].PutValue("Cherry");
            cells["B4"].PutValue("Active");
            cells["A5"].PutValue("Date");
            cells["B5"].PutValue("Inactive");

            // Hide row 3 (zero‑based index) and column B to simulate hidden cells
            sheet.Cells.Rows[2].IsHidden = true;      // hides row 3 (contains "Banana")
            sheet.Cells.Columns[1].IsHidden = true;   // hides column B (Status column)

            // Create a named range that covers A1:B5
            AsposeRange namedRange = sheet.Cells.CreateRange("A1", "B5");
            namedRange.Name = "MyRange";

            // Retrieve the named range object via the workbook's collection
            AsposeRange range = workbook.Worksheets.GetRangeByName("MyRange");

            // Build a CellArea representing the range (required by FindOptions)
            CellArea area = new CellArea
            {
                StartRow = range.FirstRow,
                StartColumn = range.FirstColumn,
                EndRow = range.FirstRow + range.RowCount - 1,
                EndColumn = range.FirstColumn + range.ColumnCount - 1
            };

            // Configure FindOptions to limit the search to the defined range
            FindOptions findOptions = new FindOptions
            {
                LookInType = LookInType.Values,
                LookAtType = LookAtType.EntireContent,
                SearchOrderByRows = true
            };
            findOptions.SetRange(area);

            // Search for the value "Active" but only consider visible cells
            string searchValue = "Active";
            Cell previous = null;
            Console.WriteLine($"Searching for \"{searchValue}\" in visible cells of named range \"MyRange\":");

            while (true)
            {
                Cell found = sheet.Cells.Find(searchValue, previous, findOptions);
                if (found == null)
                    break;

                // Verify that the cell's row and column are not hidden
                bool rowVisible = !sheet.Cells.Rows[found.Row].IsHidden;
                bool columnVisible = !sheet.Cells.Columns[found.Column].IsHidden;

                if (rowVisible && columnVisible)
                {
                    Console.WriteLine($"Found at {found.Name} (Row {found.Row + 1}, Column {found.Column + 1})");
                }

                // Continue searching after the current cell
                previous = found;
            }

            // Save the workbook
            string outputPath = "FindVisibleInNamedRange.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
