// Title: C# – Remove All Page Breaks from an Aspose.Cells Worksheet for Automatic Pagination
// Description: Load or create a Workbook, clear both HorizontalPageBreaks and VerticalPageBreaks collections, and save the file so Aspose.Cells applies automatic pagination during print or export.
// Keywords: Aspose.Cells C# page breaks | clear worksheet page breaks .NET | remove horizontal page breaks | remove vertical page breaks | automatic pagination Aspose.Cells | worksheet pagination settings | Aspose.Cells printing | Aspose.Cells PDF export | GitHub Aspose.Cells example | C# workbook pagination
// Common Searches: how to clear all page breaks in Aspose.Cells C# | remove horizontal and vertical page breaks programmatically | Aspose.Cells automatic pagination after deleting page breaks | C# code to clear page breaks in Excel workbook | Aspose.Cells page break removal before PDF export
// Developer Intent: Delete every manual horizontal and vertical page break from a worksheet so the workbook relies on automatic pagination.
// Use Cases: Prepare a template workbook for data insertion without existing page breaks interfering with print layout. | Generate a PDF or XPS from an existing Excel file with continuous pages by clearing manual breaks first. | Reset pagination in a report workbook before applying custom print settings or exporting to HTML.
// AI Prompts: Write C# code using Aspose.Cells to clear all page breaks from a specific worksheet and save the workbook. | Explain the impact of HorizontalPageBreaks.Clear() and VerticalPageBreaks.Clear() on automatic pagination in Aspose.Cells. | Provide a step‑by‑step guide to remove page breaks and then export the workbook to PDF with continuous pages.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Load or create a Workbook, clear both HorizontalPageBreaks and VerticalPageBreaks collections, and save the file so Aspose.Cells applies automatic pagination during print or export.
    public class RemoveAllPageBreaksDemo
    {
        public static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // Replace with new Workbook("input.xlsx") to load

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Remove all horizontal page breaks
            worksheet.HorizontalPageBreaks.Clear();

            // Remove all vertical page breaks
            worksheet.VerticalPageBreaks.Clear();

            // Save the workbook – automatic pagination will now be applied
            workbook.Save("Output_NoPageBreaks.xlsx");
        }
    }
}
