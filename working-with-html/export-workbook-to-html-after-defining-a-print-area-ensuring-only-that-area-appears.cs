// Title: Export a defined print area of an Excel worksheet to HTML using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export only the print area to HTML | How to set a worksheet print area before saving as HTML with Aspose.Cells | Limit HTML output to a specific range in Excel using Aspose.Cells .NET | C# code to convert cells A1:B3 to HTML with Aspose.Cells | Save workbook as HTML with a custom print area in Aspose.Cells
// Tags: Aspose.Cells export print area HTML | C# define worksheet print area | SaveFormat.Html limited range | Aspose.Cells page setup print area | HTML conversion specific Excel range

using Aspose.Cells;
using System;

// The example creates a workbook, fills cells A1:B4 with data, sets the worksheet's print area to A1:B3, and then saves the workbook as an HTML file. Only the cells within the defined print area are included in the generated HTML output.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data
        sheet.Cells["A1"].PutValue("Item");
        sheet.Cells["B1"].PutValue("Quantity");
        sheet.Cells["A2"].PutValue("Apples");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("Oranges");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue("Bananas");
        sheet.Cells["B4"].PutValue(15);

        // Define the print area (only cells A1:B3 will be exported)
        sheet.PageSetup.PrintArea = "A1:B3";

        // Export the workbook to HTML; only the defined print area appears in the output
        workbook.Save("ExportedArea.html", SaveFormat.Html);
    }
}
