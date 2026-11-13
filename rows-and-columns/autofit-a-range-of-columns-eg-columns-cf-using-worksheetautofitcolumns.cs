// Title: How to auto‑fit columns C through F in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells auto fit specific columns range example | Worksheet.AutoFitColumns with start and end column indexes in .NET | Adjust column widths for columns C to F using Aspose.Cells | Programmatically auto‑size Excel columns C-F in C# | Set column width automatically for a column range in Aspose.Cells workbook
// Tags: auto-fit column range Aspose.Cells | Worksheet.AutoFitColumns C# example | adjust Excel column widths .NET | auto-size columns by index Aspose.Cells | C# set column width for specific columns

using Aspose.Cells;
using System;

// The sample creates a new Workbook, fills cells in columns C‑F with data, calls Worksheet.AutoFitColumns(2, 5) to automatically adjust the widths of those columns, and saves the file as AutoFitColumns.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data in columns C‑F (indexes 2‑5)
        sheet.Cells["C1"].PutValue("Header C");
        sheet.Cells["D1"].PutValue("Header D");
        sheet.Cells["E1"].PutValue("Header E");
        sheet.Cells["F1"].PutValue("Header F");
        sheet.Cells["C2"].PutValue("Longer text in column C");
        sheet.Cells["D2"].PutValue("Data");
        sheet.Cells["E2"].PutValue("More data");
        sheet.Cells["F2"].PutValue("Even longer text in column F");

        // Auto‑fit columns C (index 2) to F (index 5)
        sheet.AutoFitColumns(2, 5);

        // Save the workbook
        workbook.Save("AutoFitColumns.xlsx");
    }
}
