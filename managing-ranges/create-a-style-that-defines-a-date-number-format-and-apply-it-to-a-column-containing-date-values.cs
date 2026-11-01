// Title: How to create and apply a custom mm-dd-yyyy date number format style to an entire column with Aspose.Cells for .NET (C#)
// Common Searches: how to set a custom date format for a whole column using Aspose.Cells in C# | Aspose.Cells apply mm-dd-yyyy number format to column A | C# Aspose.Cells StyleFlag to format dates in a column only | create and apply date style to Excel column with Aspose.Cells .NET | format Excel column as date without changing other cell styles using Aspose.Cells
// Tags: date number format style Aspose.Cells | column-wide style application Aspose.Cells | StyleFlag number format C# | custom mm-dd-yyyy Excel format Aspose.Cells | apply style to entire column C#

using System;
using Aspose.Cells;

// The example creates a new Workbook, writes DateTime values to cells A1‑A3, defines a style with the custom number format "mm-dd-yyyy", uses a StyleFlag to limit the applied attributes to the number format, applies this style to the whole first column, and saves the workbook as DateStyleExample.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (create rule)
        var workbook = new Workbook();

        // Get the first worksheet and its cells collection
        var sheet = workbook.Worksheets[0];
        var cells = sheet.Cells;

        // Populate column A with sample date values
        cells["A1"].PutValue(new DateTime(2023, 1, 1));
        cells["A2"].PutValue(new DateTime(2023, 2, 15));
        cells["A3"].PutValue(new DateTime(2023, 3, 30));

        // Create a style that defines a date number format (e.g., "mm-dd-yyyy")
        var dateStyle = workbook.CreateStyle();
        dateStyle.Custom = "mm-dd-yyyy";

        // Specify that only the number format should be applied
        var styleFlag = new StyleFlag();
        styleFlag.NumberFormat = true;

        // Apply the style to the entire column A (index 0)
        cells.Columns[0].ApplyStyle(dateStyle, styleFlag);

        // Save the workbook (save rule)
        workbook.Save("DateStyleExample.xlsx");
    }
}
