// Title: Set Print Area Using a Named Range in Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds sample data, defines a named range (MyPrintArea) covering A1:B3, assigns it to Worksheet.PageSetup.PrintArea, and saves as PrintAreaNamedRange.xlsx.
// Keywords: Aspose.Cells print area named range | C# set print area Aspose.Cells | Aspose.Cells PageSetup PrintArea | dynamic print region .NET | named range printing Aspose
// Common Searches: Aspose.Cells assign named range to print area C# | how to use named range for print area in .NET | set worksheet print area with name Aspose.Cells | dynamic print area example Aspose.Cells C#
// Developer Intent: Reference a named range instead of fixed cell addresses when defining a worksheet's print area.
// Use Cases: Allow users to modify the printable region by updating a named range without code changes. | Standardize print layouts across multiple sheets by applying the same named range. | Create templates where the print area adapts to varying data sizes.
// AI Prompts: Show how to change the cells covered by MyPrintArea after the workbook is created and refresh the print area. | Provide code to assign different named ranges as print areas for several worksheets in one workbook. | Explain how to read the current PrintArea value, verify it references a valid named range, and handle errors.

using System;
using Aspose.Cells;

// Creates a workbook, adds sample data, defines a named range (MyPrintArea) covering A1:B3, assigns it to Worksheet.PageSetup.PrintArea, and saves as PrintAreaNamedRange.xlsx.
class SetPrintAreaUsingNamedRange
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data
        worksheet.Cells["A1"].PutValue("Header1");
        worksheet.Cells["B1"].PutValue("Header2");
        worksheet.Cells["A2"].PutValue(10);
        worksheet.Cells["B2"].PutValue(20);
        worksheet.Cells["A3"].PutValue(30);
        worksheet.Cells["B3"].PutValue(40);

        // Define a named range that covers the area to be printed (A1:B3)
        int nameIndex = workbook.Worksheets.Names.Add("MyPrintArea");
        // The RefersTo string must start with '=' and include the sheet name
        workbook.Worksheets.Names[nameIndex].RefersTo = $"={worksheet.Name}!$A$1:$B$3";

        // Assign the named range to the worksheet's print area
        worksheet.PageSetup.PrintArea = "MyPrintArea";

        // Save the workbook
        workbook.Save("PrintAreaNamedRange.xlsx");
    }
}
