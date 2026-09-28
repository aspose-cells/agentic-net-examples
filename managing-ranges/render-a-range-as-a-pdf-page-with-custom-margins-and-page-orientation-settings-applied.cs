// Title: Export a selected worksheet range to PDF with custom margins and landscape orientation using Aspose.Cells for .NET
// AI Prompts: Generate C# code that defines a print area, sets top, bottom, left, and right margins, changes the page orientation to landscape, and saves only that range as a PDF with Aspose.Cells. | Show how to programmatically export cells A1:D10 to a PDF file while applying specific margin values and a landscape layout in Aspose.Cells for .NET.
// Common Searches: how to export a selected range to PDF with custom margins in Aspose.Cells C# | set page orientation and margins for PDF output of a range using Aspose.Cells .NET | Aspose.Cells render specific cells to PDF with landscape layout
// Tags: range PDF export using Aspose.Cells C# | configure page setup for PDF Aspose.Cells | apply landscape page orientation Aspose.Cells | define print area for PDF conversion Aspose.Cells | adjust page margins before saving PDF Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // // This example creates a workbook, fills cells A1:D10, sets that block as the print area, applies landscape orientation and custom margins, and saves the selected area as a PDF file named RangeOutput.pdf.
    class RenderRangeToPdf
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate data in the range A1:D10
                for (int row = 0; row < 10; row++)
                {
                    for (int col = 0; col < 4; col++)
                    {
                        sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }

                // Define the range to be rendered as PDF
                string rangeAddress = "A1:D10";

                // Set the print area to the defined range
                sheet.PageSetup.PrintArea = rangeAddress;

                // Set page orientation (Landscape or Portrait)
                sheet.PageSetup.Orientation = PageOrientationType.Landscape;

                // Set custom margins (values are in inches)
                sheet.PageSetup.TopMargin = 0.5;
                sheet.PageSetup.BottomMargin = 0.5;
                sheet.PageSetup.LeftMargin = 0.75;
                sheet.PageSetup.RightMargin = 0.75;

                // Output PDF file path
                string outputPath = "RangeOutput.pdf";

                // Save the workbook as PDF; only the defined print area will be exported
                workbook.Save(outputPath, SaveFormat.Pdf);
                Console.WriteLine($"PDF saved successfully to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
