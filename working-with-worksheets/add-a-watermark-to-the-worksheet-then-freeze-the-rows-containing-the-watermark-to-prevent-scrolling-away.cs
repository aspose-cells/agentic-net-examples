using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Name = "Data";

        // Populate some sample data
        for (int i = 0; i < 20; i++)
        {
            worksheet.Cells[i, 0].PutValue($"Row {i + 1}");
        }

        // Create a font for the watermark
        RenderingFont font = new RenderingFont("Arial", 48)
        {
            Bold = true,
            Color = Color.LightGray
        };

        // Create a text watermark and configure its appearance
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
        {
            HAlignment = TextAlignmentType.Center,
            VAlignment = TextAlignmentType.Center,
            Rotation = 45,
            Opacity = 0.2f,
            IsBackground = true
        };

        // Freeze the rows that contain the watermark (first two rows)
        // Freeze at cell A3 with 2 frozen rows and 0 frozen columns
        worksheet.FreezePanes("A3", 2, 0);

        // Set PDF save options with the watermark
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Watermark = watermark
        };

        // Save the workbook as PDF with the watermark applied
        workbook.Save("WatermarkedFrozen.pdf", saveOptions);
    }
}