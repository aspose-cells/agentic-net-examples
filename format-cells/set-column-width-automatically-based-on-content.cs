// Title: How to auto‑fit Excel column widths based on cell content using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells auto fit column width after adding data | Set Excel column width to fit longest text with Aspose.Cells .NET | Automatically adjust worksheet column sizes using Aspose.Cells API | Aspose.Cells AutoFitColumns example for .NET developers | Resize Excel columns to content length in C# with Aspose.Cells
// Tags: Aspose.Cells AutoFitColumns API | auto‑fit Excel columns .NET | column width auto sizing Aspose.Cells | C# Excel column width adjustment | worksheet AutoFitColumns usage

using Aspose.Cells;
using System;

// The sample creates a workbook, writes sample text to cells, calls sheet.AutoFitColumns() to automatically size each column to its longest content, and saves the file as AutoFitColumns.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data to demonstrate auto‑fit
        sheet.Cells["A1"].PutValue("Header1");
        sheet.Cells["B1"].PutValue("Header2");
        sheet.Cells["A2"].PutValue("Short");
        sheet.Cells["B2"].PutValue("This is a longer piece of text that should expand the column width");

        // Auto‑fit all columns based on their content
        sheet.AutoFitColumns();

        // Save the workbook to a file
        workbook.Save("AutoFitColumns.xlsx");
    }
}
