// Title: Load an HTML workbook, modify a cell, and save back to HTML with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells load HTML file and update cell value | How to change cell A1 in an HTML workbook using Aspose.Cells | Save modified spreadsheet to HTML with default options in Aspose.Cells .NET | Aspose.Cells example for editing HTML workbook programmatically | Convert HTML workbook to HTML after editing cells in C#
// Tags: Aspose.Cells load HTML workbook C# | Worksheet.Cells PutValue to update cell | Workbook.Save with SaveFormat.Html export | default HTML export options Aspose.Cells | programmatic cell editing in Aspose.Cells

using System;
using Aspose.Cells;

// The example loads an existing HTML file into an Aspose.Cells Workbook, updates the value of cell A1 on the first worksheet, and saves the workbook back to HTML using the default SaveFormat.Html settings.
class Program
{
    static void Main()
    {
        // Load the workbook from an existing HTML file
        Workbook workbook = new Workbook("input.html");

        // Get the first worksheet (you can change the index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Modify the value of cell A1 (change the address as required)
        worksheet.Cells["A1"].PutValue("New Value");

        // Export the workbook back to HTML using default options
        workbook.Save("output.html", SaveFormat.Html);
    }
}
