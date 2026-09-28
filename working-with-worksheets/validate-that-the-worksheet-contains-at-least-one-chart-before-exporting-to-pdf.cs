// Title: Check for at least one chart in an Excel workbook before saving as PDF with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that iterates through all worksheets in an Aspose.Cells workbook, detects if any chart objects exist, and skips the PDF export when none are found. | Show how to add a pre‑export validation step using Aspose.Cells to abort Workbook.Save when the chart collection count is zero. | Provide an example that logs a message and returns early if a workbook loaded with Aspose.Cells contains no charts before calling Save with SaveFormat.Pdf.
// Common Searches: Aspose.Cells C# verify chart presence before PDF conversion | prevent PDF export in Aspose.Cells when workbook has no charts | how to check chart count in Aspose.Cells workbook using C# | skip saving Excel to PDF with Aspose.Cells if there are no charts | C# Aspose.Cells example to abort PDF save when chart collection is empty
// Tags: Aspose.Cells chart detection before PDF export | C# workbook chart count validation | conditional PDF generation with Aspose.Cells | Excel to PDF conversion skip when no charts | pre‑export chart check Aspose.Cells

using Aspose.Cells;
using System;

// The sample loads an Excel workbook, scans each worksheet for chart objects, aborts the PDF conversion with a console message when no charts are found, and otherwise saves the workbook as a PDF using Aspose.Cells.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with actual file path)
        Workbook workbook = new Workbook("input.xlsx"); // load rule

        bool hasChart = false;

        // Check each worksheet for at least one chart
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

        // Export the workbook to PDF (replace with desired output path)
        workbook.Save("output.pdf", SaveFormat.Pdf); // save rule
    }
}
