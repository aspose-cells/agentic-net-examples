// Title: Apply 0.5‑inch custom PDF margins to all worksheets when converting an Excel file to PDF with Aspose.Cells for .NET
// AI Prompts: Write C# code that sets 0.5‑inch top, bottom, left, and right margins on every worksheet using Aspose.Cells PageSetup before saving as PDF. | Show how to set header and footer margins to zero and export the workbook to a PDF file with Aspose.Cells in a .NET application.
// Common Searches: Aspose.Cells C# set page margins before PDF export | How to define custom PDF margins for each worksheet in Aspose.Cells | Convert Excel to PDF with 0.5 inch margins using Aspose.Cells .NET | PageSetup margin values in points for PDF output with Aspose.Cells | C# Aspose.Cells PDF conversion margin configuration example
// Tags: Aspose.Cells worksheet page setup | PDF export margin settings .NET | Set page margins in points Aspose.Cells | Excel to PDF conversion margins | C# Aspose.Cells page setup API

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The program creates a workbook if missing, iterates through all worksheets, sets 0.5‑inch (36‑point) top, bottom, left, and right margins (header/footer set to zero) via the PageSetup object, and saves the workbook as a PDF using Aspose.Cells.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.pdf";

                // Ensure the input file exists; create a simple workbook if it does not.
                if (!File.Exists(inputPath))
                {
                    var tempWb = new Workbook();
                    tempWb.Worksheets[0].Cells["A1"].PutValue("Sample Data");
                    tempWb.Save(inputPath);
                }

                // Load the source workbook.
                var workbook = new Workbook(inputPath);

                // Custom margins in points (1 point = 1/72 inch).
                const double top = 36;      // 0.5 inch
                const double bottom = 36;   // 0.5 inch
                const double left = 36;     // 0.5 inch
                const double right = 36;    // 0.5 inch
                const double header = 0;
                const double footer = 0;

                // Apply margins to each worksheet.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    var pageSetup = sheet.PageSetup;
                    pageSetup.TopMargin = top;
                    pageSetup.BottomMargin = bottom;
                    pageSetup.LeftMargin = left;
                    pageSetup.RightMargin = right;
                    pageSetup.HeaderMargin = header;
                    pageSetup.FooterMargin = footer;
                }

                // Save the workbook as PDF.
                workbook.Save(outputPath, SaveFormat.Pdf);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
