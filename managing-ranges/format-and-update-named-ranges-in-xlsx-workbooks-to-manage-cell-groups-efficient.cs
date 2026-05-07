using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Saving;
using AsposeRange = Aspose.Cells.Range;

class NamedRangeManager
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate some sample data
        for (int row = 0; row < 10; row++)
        {
            for (int col = 0; col < 5; col++)
            {
                cells[row, col].PutValue($"R{row + 1}C{col + 1}");
            }
        }

        // -------------------------------------------------
        // 1. Create named ranges
        // -------------------------------------------------
        // Range A1:B5 named "FirstBlock"
        AsposeRange firstBlock = cells.CreateRange("A1", "B5");
        firstBlock.Name = "FirstBlock";

        // Range D1:E10 named "SecondBlock"
        AsposeRange secondBlock = cells.CreateRange("D1", "E10");
        secondBlock.Name = "SecondBlock";

        // Duplicate named range (to demonstrate duplicate removal)
        AsposeRange duplicate = cells.CreateRange("C1", "C3");
        duplicate.Name = "FirstBlock"; // Intentional duplicate name

        // -------------------------------------------------
        // 2. Update a named range reference
        // -------------------------------------------------
        // Change "SecondBlock" to refer to a different area (F1:G5)
        Name secondName = workbook.Worksheets.Names["SecondBlock"];
        secondName.RefersTo = $"={sheet.Name}!$F$1:$G$5";

        // -------------------------------------------------
        // 3. Apply formatting to each named range
        // -------------------------------------------------
        // Define a style with light blue background
        Style style = workbook.CreateStyle();
        style.ForegroundColor = Color.LightBlue;
        style.Pattern = BackgroundType.Solid;

        // Apply the style to "FirstBlock"
        Name firstName = workbook.Worksheets.Names["FirstBlock"];
        AsposeRange rangeToFormat = firstName.GetRange();
        rangeToFormat.ApplyStyle(style, new StyleFlag { CellShading = true });

        // Apply a different style to "SecondBlock"
        Style style2 = workbook.CreateStyle();
        style2.ForegroundColor = Color.LightGreen;
        style2.Pattern = BackgroundType.Solid;

        Name secondNameUpdated = workbook.Worksheets.Names["SecondBlock"];
        AsposeRange rangeSecond = secondNameUpdated.GetRange();
        rangeSecond.ApplyStyle(style2, new StyleFlag { CellShading = true });

        // -------------------------------------------------
        // 4. Retrieve all ranges referred by a name (example)
        // -------------------------------------------------
        // Get all ranges for "FirstBlock"
        AsposeRange[] ranges = firstName.GetRanges();
        Console.WriteLine($"\"FirstBlock\" refers to {ranges.Length} range(s).");
        foreach (AsposeRange r in ranges)
        {
            Console.WriteLine($" - Address: {r.RefersTo}");
        }

        // -------------------------------------------------
        // 5. Sort and clean up named ranges before saving
        // -------------------------------------------------
        // Sort the names collection
        workbook.Worksheets.Names.Sort();

        // Remove duplicate names (keeps the first occurrence)
        workbook.Worksheets.Names.RemoveDuplicateNames();

        // Also sort names at worksheet collection level (optional)
        workbook.Worksheets.SortNames();

        // -------------------------------------------------
        // 6. Save the workbook with SaveOptions that also sorts names
        // -------------------------------------------------
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
        {
            // Ensure names are sorted during save (redundant with earlier sorting but safe)
            SortNames = true
        };

        workbook.Save("ManagedNamedRanges.xlsx", saveOptions);

        Console.WriteLine("Workbook saved as ManagedNamedRanges.xlsx");
    }
}