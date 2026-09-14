// Title: How to enable gridlines when printing an Excel worksheet with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that sets Worksheet.PageSetup.PrintGridlines to true using Aspose.Cells and saves the workbook. | Show an example of loading an existing workbook, turning on gridline printing, and exporting it as XLSX with Aspose.Cells. | Explain how to configure page‑setup printing options to include gridlines in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# enable gridlines on printed worksheet | set PrintGridlines property before saving Excel file with Aspose.Cells | how to show Excel gridlines when printing using Aspose.Cells for .NET | page setup print options gridlines Aspose.Cells example | C# Aspose.Cells print gridlines without displaying them on screen
// Tags: Aspose.Cells print gridlines C# | worksheet page setup print options | enable gridlines on printed Excel Aspose.Cells | set PrintGridlines property Aspose.Cells | Excel workbook printing settings C#

using Aspose.Cells;
using System;

// Creates or loads a workbook, accesses the first worksheet, sets its PageSetup.PrintGridlines property to true, and saves the file, ensuring gridlines appear when the worksheet is printed.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Enable gridlines visibility when printing
        worksheet.PageSetup.PrintGridlines = true;

        // Save the workbook to a file
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
