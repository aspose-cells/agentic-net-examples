// Title: How to verify a worksheet contains charts before exporting to PDF with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# check worksheet for charts before saving as PDF | prevent PDF export when Excel sheet has no charts using Aspose.Cells | validate chart collection count in Aspose.Cells before PDF conversion | C# abort PDF generation if worksheet charts collection is empty | conditional PDF export based on chart presence in worksheet Aspose.Cells
// Tags: worksheet chart count validation Aspose.Cells | conditional PDF export Aspose.Cells C# | chart presence check before PDF conversion | export workbook to PDF with chart verification Aspose.Cells | avoid empty chart PDF generation Aspose.Cells

using Aspose.Cells;
using System;

// The example loads an Excel workbook, retrieves a worksheet, checks the Charts collection count, aborts the operation with a console message if no charts are found, and otherwise saves the workbook as a PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (or specify the desired one)
        Worksheet worksheet = workbook.Worksheets[0];

        // Validate that the worksheet contains at least one chart
        if (worksheet.Charts.Count == 0)
        {
            Console.WriteLine("The worksheet does not contain any charts. PDF export aborted.");
            return;
        }

        // Export the workbook (including the chart) to PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
        Console.WriteLine("PDF exported successfully.");
    }
}
