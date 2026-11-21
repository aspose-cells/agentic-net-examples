// Title: How to set Excel column width in pixels before inserting numeric data using Aspose.Cells for .NET (C#)
// Common Searches: aspocells set column width to 100 pixels before adding values c# | c# example using SetColumnWidthPixel to control Excel column size | how to apply pixel‑based column width in Aspose.Cells prior to putting numbers | set precise column width in Excel with Aspose.Cells then write numeric data | Aspose.Cells C# tutorial for column width pixel precision before data entry
// Tags: Aspose.Cells SetColumnWidthPixel method | C# pixel column width Excel | Excel column width control before data insertion | numeric column pixel sizing Aspose.Cells | SetColumnWidthPixel C# example

using Aspose.Cells;

// The sample creates a workbook, accesses the first worksheet, sets column B's width to 100 pixels with SetColumnWidthPixel, adds numeric values to B1‑B3, and saves the file as ColumnWidthExample.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (create rule)
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set the width of column B (index 1) to 100 pixels before adding data
        sheet.Cells.SetColumnWidthPixel(1, 100); // precise pixel width

        // Add numeric values to column B
        sheet.Cells["B1"].PutValue(123);
        sheet.Cells["B2"].PutValue(456.78);
        sheet.Cells["B3"].PutValue(0.99);

        // Save the workbook (save rule)
        workbook.Save("ColumnWidthExample.xlsx");
    }
}
