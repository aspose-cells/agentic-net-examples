// Title: Create a workbook, write values to a column, and auto‑fit the column width using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to auto‑fit a column after adding data | example code to create a workbook and auto‑size column width in Aspose.Cells .NET | Aspose.Cells C# auto‑fit column based on longest cell value | populate column A with strings and adjust width automatically using Aspose.Cells | save workbook as xlsx after auto‑fitting columns with Aspose.Cells C#
// Tags: column auto sizing Aspose.Cells C# | write values to worksheet column Aspose.Cells | initialize new workbook Aspose.Cells .NET | adjust column width based on cell content Aspose.Cells | export workbook to xlsx Aspose.Cells | auto‑fit column width API Aspose.Cells

using Aspose.Cells;
using System;

// The sample creates a new Workbook, accesses the first worksheet, writes three strings into cells A1‑A3, calls AutoFitColumn(0) to resize column A according to the longest entry, and saves the file as AutoFitColumnExample.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add data to column A (index 0)
        sheet.Cells["A1"].PutValue("Short");
        sheet.Cells["A2"].PutValue("A much longer text entry");
        sheet.Cells["A3"].PutValue("Medium length");

        // Auto‑fit column A to accommodate the longest entry
        sheet.AutoFitColumn(0);

        // Save the workbook to a file
        workbook.Save("AutoFitColumnExample.xlsx");
    }
}
