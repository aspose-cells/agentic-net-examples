// Title: Enable AutoFilter on a worksheet header row with Aspose.Cells for .NET (C#)
// Common Searches: c# asp.net how to add autofilter to first row using aspose.cells | programmatically set autofilter range for Excel sheet with Aspose.Cells | asp.net enable column sorting filter on Excel header row via Aspose.Cells | example code to apply AutoFilter to header row in a .xlsx file using C# | Aspose.Cells C# tutorial for creating Excel AutoFilter on A1:B3 range
// Tags: Aspose.Cells enable AutoFilter | C# set AutoFilter range Excel | Aspose.Cells column sorting filter | Excel header row AutoFilter C# | Aspose.Cells workbook save with AutoFilter

using Aspose.Cells;
using System;

// This example creates a new workbook, writes sample data to cells A1:B3, defines the AutoFilter range covering the header and data rows, and saves the file as AutoFilterDemo.xlsx, demonstrating how to enable column sorting via AutoFilter on the worksheet's header row using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data (optional, demonstrates the filter)
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["A2"].PutValue("Alice");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["A3"].PutValue("Bob");
        sheet.Cells["B3"].PutValue(25);

        // Enable AutoFilter on the header row (first row)
        // The range should cover all columns that need filtering; here A1:B3
        sheet.AutoFilter.Range = "A1:B3";

        // Save the workbook to a file
        workbook.Save("AutoFilterDemo.xlsx");
    }
}
