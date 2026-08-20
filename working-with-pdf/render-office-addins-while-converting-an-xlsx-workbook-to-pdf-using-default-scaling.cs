// Title: Convert XLSX to PDF with default scaling and render Office Add‑Ins using Aspose.Cells ConversionUtility (C#)
// Description: Demonstrates how to use Aspose.Cells.Utility.ConversionUtility in a .NET application to convert an XLSX workbook to PDF while preserving the workbook's original page setup. The conversion automatically renders any embedded Office Add‑Ins, and no explicit scaling options are required.
// Keywords: Aspose.Cells ConversionUtility | XLSX to PDF C# | default scaling Excel PDF | render Office Add‑Ins PDF | Aspose.Cells PDF conversion .NET | Excel page setup PDF export | C# Excel to PDF sample
// Common Searches: Aspose.Cells convert Excel to PDF default scaling | How to render Office Add‑Ins when exporting XLSX to PDF | C# ConversionUtility PDF export without scaling options | Preserve Excel page layout in PDF using Aspose.Cells | Export Excel workbook with embedded add‑ins to PDF
// Developer Intent: Convert an Excel workbook to PDF in C# while keeping the original page layout and ensuring any Office Add‑Ins are included in the PDF output.
// Use Cases: Generate PDF reports from Excel templates that contain Office Add‑Ins without manually adjusting scaling. | Batch‑process multiple spreadsheets to PDF while preserving each file’s page setup and add‑in visuals. | Distribute Excel data as PDF documents where embedded add‑ins (e.g., charts, controls) must remain visible.
// AI Prompts: Show C# code that uses Aspose.Cells ConversionUtility to convert an XLSX file to PDF with default scaling and Office Add‑Ins rendered. | Provide a .NET example for batch converting a list of Excel files to PDF, preserving page setup and logging conversion errors. | Explain how Aspose.Cells handles Office Add‑Ins during PDF export and how to ensure they appear in the final document.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Demonstrates how to use Aspose.Cells.Utility.ConversionUtility in a .NET application to convert an XLSX workbook to PDF while preserving the workbook's original page setup. The conversion automatically renders any embedded Office Add‑Ins, and no explicit scaling options are required.
class Program
{
    static void Main()
    {
        // Path to the source XLSX workbook
        string sourcePath = "input.xlsx";

        // Desired output PDF file path
        string outputPath = "output.pdf";

        // Convert the workbook to PDF using default scaling.
        // No explicit scaling options are set; the conversion uses the workbook's existing page setup.
        ConversionUtility.Convert(sourcePath, outputPath);

        Console.WriteLine("Workbook has been successfully converted to PDF.");
    }
}
