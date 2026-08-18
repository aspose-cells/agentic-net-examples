// Title: C# – Load XLSX, Apply Workbook Default Style, and Save as PDF with Aspose.Cells
// Description: Loads an existing XLSX file into a Workbook, creates a custom default style (Arial 12 pt, light‑yellow background, solid fill, text wrap), assigns it to the workbook, and exports the workbook to a PDF using SaveFormat.Pdf.
// Keywords: Aspose.Cells XLSX to PDF | C# workbook default style | Aspose.Cells custom style PDF conversion | SaveFormat.Pdf Aspose.Cells | Excel to PDF with formatting .NET | Aspose.Cells style example GitHub | global Aspose.Cells PDF export
// Common Searches: How to set a default style for an entire workbook before PDF conversion in Aspose.Cells | C# code to load an XLSX file and export it as PDF with Aspose.Cells | Apply custom formatting to all cells and save as PDF using Aspose.Cells .NET | Aspose.Cells example for workbook style and PDF output
// Developer Intent: The developer needs to read an XLSX workbook, apply a uniform custom style to the whole workbook, and generate a PDF version of the spreadsheet.
// Use Cases: Brand a spreadsheet with company colors and fonts before creating a printable PDF report. | Produce a PDF that preserves text wrapping and highlighted backgrounds for easier reading. | Standardize formatting across multiple workbooks in a batch conversion pipeline.
// AI Prompts: Write C# code that opens an Excel file, sets a custom default style for the workbook, and saves it as a PDF using Aspose.Cells. | Explain how to extend the default style to include borders, alignment, and number formats before PDF export. | Provide a step‑by‑step guide to apply different styles to specific ranges and then convert the workbook to PDF with Aspose.Cells.

using System;
using System.Drawing;                     // For Color
using Aspose.Cells;                       // Core Aspose.Cells namespace

namespace AsposeCellsPdfConversion
{
    // Loads an existing XLSX file into a Workbook, creates a custom default style (Arial 12 pt, light‑yellow background, solid fill, text wrap), assigns it to the workbook, and exports the workbook to a PDF using SaveFormat.Pdf.
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Path for the resulting PDF file
            string destPath = "output.pdf";

            // Load the workbook from the existing XLSX file
            // Uses the Workbook(string) constructor (lifecycle rule)
            Workbook workbook = new Workbook(sourcePath);

            // -------------------------------------------------
            // Create and configure a custom style
            // -------------------------------------------------
            // Create a new style object (CreateStyle rule)
            Style customStyle = workbook.CreateStyle();

            // Example customizations
            customStyle.Font.Name = "Arial";
            customStyle.Font.Size = 12;
            customStyle.ForegroundColor = Color.LightYellow;
            customStyle.Pattern = BackgroundType.Solid;
            customStyle.IsTextWrapped = true;

            // Apply the custom style to the default style of the workbook
            // (you could also apply it to specific cells/ranges if needed)
            workbook.DefaultStyle = customStyle;

            // -------------------------------------------------
            // Save the workbook as a PDF document
            // -------------------------------------------------
            // Uses the Save(string, SaveFormat) method (lifecycle rule)
            workbook.Save(destPath, SaveFormat.Pdf);

            Console.WriteLine("Workbook has been converted to PDF successfully.");
        }
    }
}
