// Title: Freeze the first two columns (A and B) in an Excel worksheet with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to lock columns A and B while allowing all rows to scroll. | Show how to invoke Worksheet.FreezePanes to apply a column‑only freeze in a newly created workbook and export it as an Xlsx file. | Provide a concise example that creates a worksheet, populates header cells, and freezes the first two columns with Aspose.Cells.
// Common Searches: Aspose.Cells C# freeze only the first two columns in an Excel workbook | How to apply FreezePanes to lock columns A and B with Aspose.Cells for .NET | Prevent horizontal scrolling of identifier columns in Excel using Aspose.Cells | C# example for column‑only freeze panes with Aspose.Cells | Freeze columns without freezing rows Aspose.Cells .NET
// Tags: Aspose.Cells FreezePanes column only | C# freeze first two Excel columns | Aspose.Cells worksheet column freeze example | Excel Xlsx column freeze Aspose.Cells | prevent horizontal scrolling identifier columns Aspose.Cells

using System;
using Aspose.Cells;

// Creates a new Workbook, adds header values to cells A1‑D1, freezes columns A and B using Worksheet.FreezePanes(0, 2, 0, 0) (rows remain unfrozen), and saves the file as FrozenColumns.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Optional: add some sample data
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Code");
            sheet.Cells["C1"].PutValue("Name");
            sheet.Cells["D1"].PutValue("Description");

            // Freeze the first two columns (A and B)
            // Parameters: firstRow, firstColumn, totalRows, totalColumns
            // Setting firstRow and totalRows to 0 means no rows are frozen.
            // Setting firstColumn to 2 freezes columns A and B.
            sheet.FreezePanes(0, 2, 0, 0);

            // Save the workbook
            workbook.Save("FrozenColumns.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
