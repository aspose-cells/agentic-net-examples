// Title: C# – Convert an Excel Table to a Range (first 10 rows) while preserving formatting with Aspose.Cells for .NET
// Description: Loads a workbook, selects the first ListObject, and uses TableToRangeOptions (LastRow = 9) to turn the table into a normal range that includes only rows 0‑9. All original table styles are kept and the workbook is saved as a new file. Ideal as a concise Aspose.Cells for .NET sample.
// Keywords: Aspose.Cells | C# | Convert Excel table to range | TableToRangeOptions | LastRow | preserve formatting | first 10 rows | ListObject conversion | sample code | GitHub example | API usage
// Common Searches: Aspose.Cells convert table to range C# | limit table conversion to first 10 rows Aspose | preserve Excel table formatting when converting to range | TableToRangeOptions LastRow example | how to turn ListObject into range using Aspose.Cells
// Developer Intent: Convert a specific worksheet table into a plain range limited to the first ten rows while retaining its visual style.
// Use Cases: Export data to systems that only accept plain ranges, not structured tables. | Create a quick preview sheet that shows only the top ten rows of a large table with original styling. | Prepare workbooks for legacy applications that cannot process Excel tables.
// AI Prompts: Write C# code with Aspose.Cells to convert the second ListObject in a worksheet to a range limited to row 15, keeping all formatting. | Explain the role of TableToRangeOptions.LastRow and how it trims rows when converting a table to a range in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Loads a workbook, selects the first ListObject, and uses TableToRangeOptions (LastRow = 9) to turn the table into a normal range that includes only rows 0‑9. All original table styles are kept and the workbook is saved as a new file. Ideal as a concise Aspose.Cells for .NET sample.
class ConvertTableToRange
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Assume the table we want to convert is the first ListObject in the sheet
        if (sheet.ListObjects.Count == 0)
        {
            Console.WriteLine("No tables found in the worksheet.");
            return;
        }

        ListObject table = sheet.ListObjects[0];

        // Create conversion options to limit conversion up to the 10th row (zero‑based index 9)
        TableToRangeOptions options = new TableToRangeOptions
        {
            LastRow = 9   // rows 0‑9 will be kept as a range; rows beyond are excluded
        };

        // Convert the table to a normal range while preserving its formatting
        table.ConvertToRange(options);

        // Save the modified workbook (replace with desired output path)
        workbook.Save("output.xlsx");

        Console.WriteLine("Table successfully converted to range up to row 10 and workbook saved.");
    }
}
