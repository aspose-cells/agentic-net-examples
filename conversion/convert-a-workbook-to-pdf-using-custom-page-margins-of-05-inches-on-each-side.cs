// Title: Export an Aspise.Cells Workbook to PDF with 0.5‑inch margins (C#)
// Description: Demonstrates how to set left, right, top, and bottom margins to 0.5 inches using the PageSetup properties of a worksheet, then save the workbook as a PDF via SaveFormat.Pdf. Includes optional sample data to illustrate the margin effect.
// Keywords: Aspose.Cells PDF margins C# | custom page margins Aspose.Cells | Workbook.Save PDF custom margins | set margins inches Aspose.Cells | export Excel to PDF half‑inch margins | C# Aspose.Cells page setup | PDF export with uniform margins
// Common Searches: Aspose.Cells set 0.5 inch margins for PDF export | C# export Excel workbook to PDF with custom margins | How to change page margins in Aspose.Cells before saving as PDF | Half‑inch margin PDF using Aspose.Cells C# | PageSetup margin properties Aspose.Cells example
// Developer Intent: The developer needs to generate a PDF from an Excel workbook while enforcing a consistent 0.5‑inch margin on every side of each page.
// Use Cases: Printing reports that must adhere to a half‑inch printable area. | Creating PDF invoices or contracts where precise margin specifications are required. | Automating documentation generation from Excel templates with uniform margins for binding or filing.
// AI Prompts: Generate C# code that uses Aspose.Cells to export a workbook to PDF with 0.5‑inch margins on all sides. | Explain the impact of PageSetup.LeftMarginInch, RightMarginInch, TopMarginInch, and BottomMarginInch on PDF layout in Aspose.Cells. | Show how to convert margin values from centimeters to inches and apply them before saving a workbook as PDF with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsMarginPdfDemo
{
    // Demonstrates how to set left, right, top, and bottom margins to 0.5 inches using the PageSetup properties of a worksheet, then save the workbook as a PDF via SaveFormat.Pdf. Includes optional sample data to illustrate the margin effect.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set custom page margins (0.5 inches on each side)
            sheet.PageSetup.LeftMarginInch = 0.5;
            sheet.PageSetup.RightMarginInch = 0.5;
            sheet.PageSetup.TopMarginInch = 0.5;
            sheet.PageSetup.BottomMarginInch = 0.5;

            // (Optional) Add some data to visualize the margins in the PDF
            sheet.Cells["A1"].PutValue("Demo of 0.5 inch margins");
            sheet.Cells["A2"].PutValue("Each side of the page has a half‑inch margin.");

            // Save the workbook as PDF
            string outputPath = "WorkbookWithMargins.pdf";
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"Workbook saved to PDF with custom margins: {outputPath}");
        }
    }
}
