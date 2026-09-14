// Title: Remove all pivot tables from an Excel workbook and save the cleaned report as a PDF with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file, deletes every pivot table in all worksheets using Aspose.Cells, and then saves the result as a PDF. | Show how to use the Worksheets.ClearPivottables method in Aspose.Cells to clean a workbook before exporting it to PDF. | Provide a step‑by‑step example that programmatically clears pivot tables from a workbook and performs PDF conversion with Aspose.Cells for .NET.
// Common Searches: aspnet remove pivot tables from Excel workbook before PDF export | c# Aspose.Cells clear all pivot tables and convert to PDF | how to delete pivot tables in every sheet using Aspose.Cells | export cleaned Excel file to PDF after removing pivot tables with Aspose.Cells | Aspose.Cells Worksheets.ClearPivottables usage example
// Tags: clear pivot tables Aspose.Cells | Worksheets.ClearPivottables method | export workbook to PDF Aspose.Cells | remove pivot tables C# | PDF conversion after pivot cleanup

using System;
using Aspose.Cells;

// The code loads 'input.xlsx', calls Worksheets.ClearPivottables() to delete all pivot tables across every worksheet, and then saves the cleaned workbook as 'clean_report.pdf' using the PDF save format.
class Program
{
    static void Main()
    {
        // Load the existing workbook that contains pivot tables
        Workbook workbook = new Workbook("input.xlsx");

        // Remove all pivot tables from every worksheet to clean the report
        workbook.Worksheets.ClearPivottables();

        // Save the cleaned workbook as a PDF document
        workbook.Save("clean_report.pdf", SaveFormat.Pdf);
    }
}
