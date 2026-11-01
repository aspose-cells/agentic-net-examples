// Title: Check for charts in an Excel workbook before exporting to PDF with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells how to detect if a worksheet contains any charts before saving as PDF | prevent PDF generation in Aspose.Cells when workbook has no chart objects | verify chart presence in Excel file using Aspose.Cells .NET prior to PDF export | skip PDF conversion if workbook worksheets have zero charts Aspose.Cells C# | conditional PDF save based on chart count in Aspose.Cells workbook
// Tags: chart detection before PDF export Aspose.Cells | conditional workbook.Save to PDF based on chart count | worksheet chart count validation C# | Aspose.Cells abort PDF when no charts | Excel to PDF conversion with chart prerequisite

using Aspose.Cells;
using System;

// The example loads an Excel workbook, scans all worksheets for at least one chart, aborts the PDF conversion with a console message if none are found, and otherwise saves the workbook as a PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the workbook (load rule)
        var workbook = new Workbook("input.xlsx");

        // Check if any worksheet contains at least one chart
        bool hasChart = false;
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            if (sheet.Charts.Count > 0)
            {
                hasChart = true;
                break;
            }
        }

        if (!hasChart)
        {
            Console.WriteLine("No charts found in the workbook. PDF export aborted.");
            return;
        }

        // Export the workbook to PDF (save rule)
        workbook.Save("output.pdf", SaveFormat.Pdf);
        Console.WriteLine("Workbook exported to PDF successfully.");
    }
}
