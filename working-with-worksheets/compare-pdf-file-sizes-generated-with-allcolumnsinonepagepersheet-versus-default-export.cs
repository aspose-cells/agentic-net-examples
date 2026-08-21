// Title: Aspose.Cells C# – Compare PDF file size with AllColumnsInOnePagePerSheet vs default export
// Description: This C# example creates a workbook with 50 columns, saves it twice as PDF—once using the default settings and once with PdfSaveOptions.AllColumnsInOnePagePerSheet = true (and OnePagePerSheet = true)—then reads the file sizes and reports which setting yields a smaller PDF.
// Keywords: Aspose.Cells | C# | .NET | PdfSaveOptions | AllColumnsInOnePagePerSheet | OnePagePerSheet | PDF export size | file size comparison | worksheet to PDF | Aspose.Cells PDF options
// Common Searches: AllColumnsInOnePagePerSheet PDF size Aspose.Cells | default PDF export size Aspose.Cells .NET | compare PDF file size Aspose.Cells | does OnePagePerSheet affect PDF size | Aspose.Cells PDF file size optimization
// Developer Intent: Determine whether enabling AllColumnsInOnePagePerSheet (with OnePagePerSheet) produces a larger or smaller PDF than the default export.
// Use Cases: Measure baseline PDF size for a wide worksheet. | Evaluate storage impact of AllColumnsInOnePagePerSheet before selecting export settings. | Automate size comparison in CI pipelines to enforce PDF size limits. | Show PDF rendering differences for documentation or reporting.
// AI Prompts: Generate C# code that calculates and prints the percentage difference between the default PDF size and the AllColumnsInOnePagePerSheet PDF using Aspose.Cells. | Explain how AllColumnsInOnePagePerSheet interacts with OnePagePerSheet and its effect on PDF rendering speed and file size in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsComparison
{
    // This C# example creates a workbook with 50 columns, saves it twice as PDF—once using the default settings and once with PdfSaveOptions.AllColumnsInOnePagePerSheet = true (and OnePagePerSheet = true)—then reads the file sizes and reports which setting yields a smaller PDF.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add sample data that spans many columns
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate 50 columns with sample data to make the width noticeable
            for (int col = 0; col < 50; col++)
            {
                sheet.Cells[0, col].PutValue("Header " + (col + 1));
                sheet.Cells[1, col].PutValue("Data " + (col + 1));
            }

            // -----------------------------------------------------------------
            // Save PDF with default settings (no AllColumnsInOnePagePerSheet)
            // -----------------------------------------------------------------
            string defaultPdfPath = "DefaultExport.pdf";
            workbook.Save(defaultPdfPath, new PdfSaveOptions());

            // -----------------------------------------------------------------
            // Save PDF with AllColumnsInOnePagePerSheet = true
            // -----------------------------------------------------------------
            string allColumnsPdfPath = "AllColumnsOnePage.pdf";
            PdfSaveOptions allColumnsOptions = new PdfSaveOptions
            {
                AllColumnsInOnePagePerSheet = true,
                OnePagePerSheet = true // ensure content fits on a single page per sheet
            };
            workbook.Save(allColumnsPdfPath, allColumnsOptions);

            // -----------------------------------------------------------------
            // Compare file sizes
            // -----------------------------------------------------------------
            long defaultSize = new FileInfo(defaultPdfPath).Length;
            long allColumnsSize = new FileInfo(allColumnsPdfPath).Length;

            Console.WriteLine($"Default PDF size: {defaultSize} bytes");
            Console.WriteLine($"AllColumnsInOnePagePerSheet PDF size: {allColumnsSize} bytes");

            if (allColumnsSize < defaultSize)
                Console.WriteLine("AllColumnsInOnePagePerSheet produced a smaller PDF.");
            else if (allColumnsSize > defaultSize)
                Console.WriteLine("AllColumnsInOnePagePerSheet produced a larger PDF.");
            else
                Console.WriteLine("Both PDFs have the same size.");
        }
    }
}
