using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsThemePdfPreview
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Prepare 12 custom theme colors (required length)
            Color[] customColors = new Color[]
            {
                Color.FromArgb(255, 255, 255), // Background1
                Color.FromArgb(0, 0, 0),       // Text1
                Color.FromArgb(240, 240, 240), // Background2
                Color.FromArgb(50, 50, 50),    // Text2
                Color.FromArgb(0, 120, 215),   // Accent1
                Color.FromArgb(232, 17, 35),   // Accent2
                Color.FromArgb(255, 185, 0),   // Accent3
                Color.FromArgb(0, 204, 106),   // Accent4
                Color.FromArgb(0, 153, 188),   // Accent5
                Color.FromArgb(255, 102, 0),   // Accent6
                Color.FromArgb(0, 0, 255),     // Hyperlink
                Color.FromArgb(128, 0, 128)    // Followed Hyperlink
            };

            // Apply the custom theme to the workbook
            workbook.CustomTheme("MyCustomTheme", customColors);

            // Hide all worksheets except the first one to generate a preview of the first sheet only
            for (int i = 1; i < workbook.Worksheets.Count; i++)
            {
                workbook.Worksheets[i].IsVisible = false;
            }

            // Save the first sheet as a PDF preview
            string outputPdf = "FirstSheetPreview.pdf";
            workbook.Save(outputPdf, SaveFormat.Pdf);

            Console.WriteLine($"PDF preview generated: {outputPdf}");
        }
    }
}