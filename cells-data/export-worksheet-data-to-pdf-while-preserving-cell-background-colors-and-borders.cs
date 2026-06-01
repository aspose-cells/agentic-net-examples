using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExportWorksheetToPdfWithFormatting
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Price");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["B2"].PutValue(1.20);
            worksheet.Cells["A3"].PutValue("Banana");
            worksheet.Cells["B3"].PutValue(0.80);

            // Apply background color and borders to header row
            Style headerStyle = workbook.CreateStyle();
            headerStyle.ForegroundColor = Color.LightBlue;
            headerStyle.Pattern = BackgroundType.Solid;
            headerStyle.Font.IsBold = true; // set bold font
            headerStyle.SetBorder(BorderType.BottomBorder, CellBorderType.Thick, Color.DarkBlue);
            headerStyle.SetBorder(BorderType.TopBorder, CellBorderType.Thick, Color.DarkBlue);
            headerStyle.SetBorder(BorderType.LeftBorder, CellBorderType.Thick, Color.DarkBlue);
            headerStyle.SetBorder(BorderType.RightBorder, CellBorderType.Thick, Color.DarkBlue);
            worksheet.Cells.CreateRange("A1:B1").ApplyStyle(headerStyle, new StyleFlag { All = true });

            // Apply background color and borders to data cells
            Style dataStyle = workbook.CreateStyle();
            dataStyle.ForegroundColor = Color.LightYellow;
            dataStyle.Pattern = BackgroundType.Solid;
            dataStyle.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Gray);
            dataStyle.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Gray);
            dataStyle.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Gray);
            dataStyle.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Gray);
            worksheet.Cells.CreateRange("A2:B3").ApplyStyle(dataStyle, new StyleFlag { All = true });

            // PDF save options with document structure export (preserves formatting)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true
            };

            string outputPath = "WorksheetWithFormatting.pdf";

            // Save the workbook as PDF
            try
            {
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine("PDF exported successfully with background colors and borders preserved.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save PDF: {saveEx.Message}");
                throw;
            }
        }
    }
}