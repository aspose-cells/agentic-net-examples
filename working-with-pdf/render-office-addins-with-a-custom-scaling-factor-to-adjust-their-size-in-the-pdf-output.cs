using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Saving;

namespace OfficeAddInPdfScalingDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Populate some sample data (simulating an Office Add‑In content)
            sheet.Cells["A1"].PutValue("Office Add‑In Demo");
            sheet.Cells["A2"].PutValue("This content will be scaled in the PDF output.");
            for (int i = 3; i <= 12; i++)
            {
                sheet.Cells[$"A{i}"].PutValue($"Item {i - 2}");
                sheet.Cells[$"B{i}"].PutValue((i - 2) * 10);
            }

            // 3. (Optional) Add a shape to represent an Add‑In UI element
            sheet.Shapes.AddRectangle(2, 2, 100, 50, 200, 100);

            // 4. Configure page scaling:
            //    - Set a custom zoom factor (e.g., 150%).
            //    - Ensure the scaling is based on the percent value.
            sheet.PageSetup.Zoom = 150;               // custom scaling factor
            sheet.PageSetup.IsPercentScale = true;    // use percent scaling

            // 5. Use SheetRender to obtain the calculated page scale (for verification)
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions();
            SheetRender renderer = new SheetRender(sheet, renderOptions);
            double calculatedScale = renderer.PageScale; // e.g., 1.5 for 150%
            Console.WriteLine($"Calculated page scale (factor): {calculatedScale}");

            // 6. Save the workbook as PDF with the applied scaling
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            // No explicit scaling property on PdfSaveOptions; the worksheet's PageSetup.Zoom is respected.
            workbook.Save("OfficeAddInScaled.pdf", pdfOptions);

            // Clean up
            renderer.Dispose();

            Console.WriteLine("PDF saved with custom scaling factor.");
        }
    }
}