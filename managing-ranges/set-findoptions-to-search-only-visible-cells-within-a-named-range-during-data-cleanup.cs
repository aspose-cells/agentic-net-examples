using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class SearchVisibleInNamedRange
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (some rows are hidden)
                sheet.Cells["A1"].PutValue("Apple");
                sheet.Cells["A2"].PutValue("Banana");
                sheet.Cells["A3"].PutValue("Cherry");
                sheet.Cells["A4"].PutValue("Date");
                sheet.Cells["A5"].PutValue("Elderberry");

                // Hide row 3 (zero‑based index) so it will be invisible during the search
                sheet.Cells.HideRow(2); // Row 3 is hidden

                // Create a named range that covers A1:A5
                Aspose.Cells.Range range = sheet.Cells.CreateRange("A1", "A5");
                range.Name = "FruitRange";

                // Retrieve the named range object via its name
                Aspose.Cells.Range namedRange = workbook.Worksheets.GetRangeByName("FruitRange");

                // Build a CellArea representing the range (required by FindOptions)
                CellArea area = new CellArea
                {
                    StartRow = namedRange.FirstRow,
                    StartColumn = namedRange.FirstColumn,
                    EndRow = namedRange.FirstRow + namedRange.RowCount - 1,
                    EndColumn = namedRange.FirstColumn + namedRange.ColumnCount - 1
                };

                // Configure FindOptions (search only within the defined CellArea)
                FindOptions options = new FindOptions
                {
                    LookInType = LookInType.Values,
                    LookAtType = LookAtType.EntireContent
                };
                options.SetRange(area);

                // Perform the find operation for the value "Cherry"
                Cell found = sheet.Cells.Find("Cherry", null, options);

                // If the found cell is in a hidden row, treat it as not found
                if (found != null && sheet.Cells.IsRowHidden(found.Row))
                {
                    found = null;
                }

                Console.WriteLine(found != null
                    ? $"Found '{found.StringValue}' at {found.Name}"
                    : "Value not found in visible cells.");

                // Save the workbook
                string outputPath = "SearchVisibleInNamedRange.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}