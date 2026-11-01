// Title: Enable gridlines when printing a worksheet with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, sets PageSetup.PrintGridlines to true for the first worksheet, and saves it as an XLSX file. | Show a minimal Aspose.Cells example that configures printing options to include gridlines in the output document. | Demonstrate how to toggle the PrintGridlines flag on a worksheet's PageSetup and persist the changes with Aspose.Cells.
// Common Searches: Aspose.Cells C# print worksheet with gridlines visible | set PrintGridlines property in Aspose.Cells .NET example | enable gridlines for printed Excel file using Aspose.Cells API | how to show Excel gridlines when exporting to XLSX with Aspose.Cells | C# Aspose.Cells page setup print gridlines configuration
// Tags: Aspose.Cells worksheet PrintGridlines setting | C# enable printing gridlines Aspose.Cells | Aspose.Cells page setup print options | export Excel with gridlines using Aspose.Cells | save workbook with visible gridlines .NET

using Aspose.Cells;
using System;

// // Creates a new workbook, turns on PrintGridlines for the first worksheet, and saves it as PrintedWithGridlines.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Enable gridlines visibility when printing
        sheet.PageSetup.PrintGridlines = true;

        // Save the workbook
        workbook.Save("PrintedWithGridlines.xlsx", SaveFormat.Xlsx);
    }
}
