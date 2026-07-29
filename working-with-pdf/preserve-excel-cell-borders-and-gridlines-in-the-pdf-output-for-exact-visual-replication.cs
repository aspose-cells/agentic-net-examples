// Title: C# – Preserve Excel Cell Borders and Gridlines When Converting to PDF with Aspose.Cells
// Description: A .NET example that creates a workbook, fills a small data range, applies a thick blue border to each used cell, makes gridlines visible and printable, sets PdfSaveOptions (dotted gray gridlines), and saves the worksheet as a PDF that exactly replicates the original borders and gridlines.
// Keywords: Aspose.Cells C# PDF conversion | preserve cell borders PDF | include gridlines Aspose.Cells | PdfSaveOptions GridlineType | Excel to PDF exact visual replica | Aspose.Cells .NET export | border style PDF Aspose | gridline color PDF | C# Excel PDF export | Aspose.Cells USA | Aspose.Cells Europe
// Common Searches: How to keep Excel borders in PDF using Aspose.Cells C# | Aspose.Cells PDFSaveOptions gridlines | Export worksheet with thick borders to PDF .NET | Preserve gridlines when saving Excel as PDF Aspose | C# code to apply borders to used range and export PDF | Aspose.Cells PDF export settings for exact layout
// Developer Intent: Generate a PDF from an Excel workbook that faithfully reproduces the worksheet’s custom cell borders and visible gridlines.
// Use Cases: Print-ready reports where the PDF must match the on‑screen Excel layout, including custom borders and dotted gridlines. | Invoice generation from a template workbook that requires retained border styling for regulatory compliance. | Multi‑page data sheets exported as PDFs where gridlines aid readability and data interpretation.
// AI Prompts: Show me how to change the border color and line style while still preserving gridlines in the PDF output using Aspose.Cells. | Provide an example that applies borders to a dynamic used range and exports the worksheet to PDF with gridlines in C#. | Explain how to configure PdfSaveOptions to hide gridlines or adjust their opacity for different PDF export scenarios.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// A .NET example that creates a workbook, fills a small data range, applies a thick blue border to each used cell, makes gridlines visible and printable, sets PdfSaveOptions (dotted gray gridlines), and saves the worksheet as a PDF that exactly replicates the original borders and gridlines.
class PreserveBordersAndGridlinesPdf
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data
            worksheet.Cells["A1"].PutValue("Header");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("Item1");
            worksheet.Cells["B2"].PutValue(100);
            worksheet.Cells["A3"].PutValue("Item2");
            worksheet.Cells["B3"].PutValue(200);

            // Define a thick blue border style
            Style borderStyle = workbook.CreateStyle();
            borderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thick;
            borderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
            borderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thick;
            borderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thick;
            borderStyle.Borders[BorderType.TopBorder].Color = Color.Blue;
            borderStyle.Borders[BorderType.BottomBorder].Color = Color.Blue;
            borderStyle.Borders[BorderType.LeftBorder].Color = Color.Blue;
            borderStyle.Borders[BorderType.RightBorder].Color = Color.Blue;

            // Apply the border style to the used range manually (avoids ApplyStyle overload issue)
            CellArea usedRange = new CellArea { StartRow = 0, StartColumn = 0, EndRow = 2, EndColumn = 1 };
            for (int row = usedRange.StartRow; row <= usedRange.EndRow; row++)
            {
                for (int col = usedRange.StartColumn; col <= usedRange.EndColumn; col++)
                {
                    worksheet.Cells[row, col].SetStyle(borderStyle);
                }
            }

            // Make gridlines visible in the worksheet and ensure they are printed
            worksheet.IsGridlinesVisible = true;
            worksheet.PageSetup.PrintGridlines = true;

            // Configure PDF save options to include gridlines with desired appearance
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                GridlineType = GridlineType.Dotted,   // Choose gridline style
                GridlineColor = Color.Gray            // Choose gridline color
            };

            // Save the workbook as PDF preserving borders and gridlines
            workbook.Save("PreservedBordersGridlines.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
