using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class ScalingDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate the worksheet with sample data
        for (int i = 0; i < 5; i++)
        {
            sheet.Cells[i, 0].PutValue($"Item {i + 1}");
            sheet.Cells[i, 1].PutValue((i + 1) * 10);
        }

        // Define a scaling factor using PageSetup.Zoom (percentage)
        // This scales the worksheet content when rendered or printed
        sheet.PageSetup.Zoom = 150; // 150% scaling

        // Configure image rendering options with a desired size
        // keepAspectRatio set to false ensures proportional scaling based on the size we provide
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.OnePagePerSheet = true;
        options.SetDesiredSize(1200, 800, false); // width, height, no aspect ratio lock

        // Render the worksheet to an image using the scaling settings
        SheetRender renderer = new SheetRender(sheet, options);
        renderer.ToImage(0, "scaled_output.png");

        // Save the workbook (the Zoom setting is persisted in the file)
        workbook.Save("scaled_workbook.xlsx");
    }
}