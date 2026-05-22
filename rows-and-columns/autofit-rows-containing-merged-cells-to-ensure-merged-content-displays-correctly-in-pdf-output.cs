using System;
using Aspose.Cells;

namespace AutoFitMergedRowsPdfDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Put a long text into the top‑left cell
            worksheet.Cells["A1"].PutValue("This is a long piece of text that should be displayed correctly after auto‑fitting rows that contain merged cells. It spans multiple lines when wrapped.");

            // Enable text wrapping for the cell so the content can occupy several lines
            Style style = worksheet.Cells["A1"].GetStyle();
            style.IsTextWrapped = true;
            worksheet.Cells["A1"].SetStyle(style);

            // Merge a range of cells (A1:C3) to simulate a merged area
            worksheet.Cells.Merge(0, 0, 3, 3); // rows 0‑2, columns 0‑2

            // Configure auto‑fitter options to expand each row of the merged area
            AutoFitterOptions options = new AutoFitterOptions
            {
                AutoFitMergedCellsType = AutoFitMergedCellsType.EachLine,
                // Optional: ensure wrapped text is considered as a paragraph
                AutoFitWrappedTextType = AutoFitWrappedTextType.Paragraph
            };

            // Auto‑fit rows using the specified options
            worksheet.AutoFitRows(options);

            // Save the workbook as PDF (the auto‑fitted row heights are reflected in the PDF)
            workbook.Save("AutoFitMergedRows.pdf", SaveFormat.Pdf);
        }
    }
}