// Title: Auto‑fit a single column in an Excel worksheet with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells auto fit column width based on cell values | using Worksheet.AutoFitColumn to resize column A in a .NET Excel workbook | example code for auto‑sizing a specific column with Aspose.Cells | how to automatically adjust column width after adding data in Aspose.Cells C# | auto fit a single column in an Excel sheet using Aspose.Cells .NET
// Tags: Worksheet.AutoFitColumn method C# | adjust Excel column width programmatically .NET | single column auto sizing Aspose.Cells | column width adjustment based on cell content | auto fit column programmatically Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a new Workbook, writes sample text into cells A1‑A3 of the first worksheet, calls Worksheet.AutoFitColumn(0) to automatically size column A according to its content, and saves the file as AutoFitColumn.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data to column A
        sheet.Cells["A1"].PutValue("Header");
        sheet.Cells["A2"].PutValue("Short");
        sheet.Cells["A3"].PutValue("This is a much longer text that requires auto‑fit");

        // Auto‑fit column 0 (column A) based on its content
        sheet.AutoFitColumn(0);

        // Save the workbook
        workbook.Save("AutoFitColumn.xlsx");
    }
}
