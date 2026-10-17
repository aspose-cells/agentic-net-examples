// Title: Convert an Excel workbook to PDF and set the workbook’s CreatedTime to the current timestamp using Aspose.Cells for .NET (C#)
// Common Searches: c# aspose.cells convert xlsx to pdf and set createdtime to now | how to update Excel workbook BuiltInDocumentProperties.CreatedTime before exporting to PDF in .NET | set audit creation date on Excel file using Aspose.Cells then save as PDF | Aspose.Cells example for saving workbook as PDF with current timestamp metadata | C# code to change Excel file creation date and generate PDF with Aspose.Cells
// Tags: aspose.cells set createdtime property c# | excel to pdf conversion with aspose.cells .net | audit metadata update before pdf export c# | builtindocumentproperties usage in aspose.cells | save workbook as pdf with timestamp c# | c# aspose.cells workbook conversion pdf

using System;
using Aspose.Cells;

// The sample loads an Excel workbook, assigns DateTime.Now to its BuiltInDocumentProperties.CreatedTime for audit tracking, and then saves the workbook as a PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Set the creation date to the current timestamp for audit tracking
        workbook.BuiltInDocumentProperties.CreatedTime = DateTime.Now;

        // Convert and save the workbook as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
