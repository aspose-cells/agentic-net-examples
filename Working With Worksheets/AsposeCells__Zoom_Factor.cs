using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set worksheet view zoom (percentage)
        sheet.Zoom = 120; // 120%

        // Set page setup zoom (print scaling) and enable percent scale
        sheet.PageSetup.Zoom = 150; // 150%
        sheet.PageSetup.IsPercentScale = true;

        // Output current zoom settings
        Console.WriteLine($"Worksheet view zoom: {sheet.Zoom}%");
        Console.WriteLine($"PageSetup zoom: {sheet.PageSetup.Zoom}%");

        // Save workbook with OoxmlSaveOptions that updates zoom before saving
        OoxmlSaveOptions ooxmlOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
        {
            UpdateZoom = true
        };
        string xlsxPath = "ZoomDemo.xlsx";
        workbook.Save(xlsxPath, ooxmlOptions);
        Console.WriteLine($"Workbook saved to {xlsxPath} with UpdateZoom enabled.");

        // Save workbook to HTML with worksheet scalable option (honors Worksheet.Zoom)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            WorksheetScalable = true
        };
        string htmlPath = "ZoomDemo.html";
        workbook.Save(htmlPath, htmlOptions);
        Console.WriteLine($"Workbook saved to {htmlPath} with WorksheetScalable enabled.");

        // Render the first page to an image and display the calculated page scale
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
        {
            ImageType = Aspose.Cells.Drawing.ImageType.Png
        };
        SheetRender renderer = new SheetRender(sheet, imgOptions);
        double pageScale = renderer.PageScale;
        Console.WriteLine($"Calculated page scale from SheetRender: {pageScale}");

        string imgPath = "ZoomDemo.png";
        renderer.ToImage(0, imgPath);
        Console.WriteLine($"Rendered sheet page saved as image to {imgPath}.");

        // Demonstrate DPI effect on saved file
        double originalDpi = CellsHelper.DPI;
        Console.WriteLine($"Original DPI: {originalDpi}");
        CellsHelper.DPI = 300;
        string dpiPath = "ZoomDemo_DPI300.xlsx";
        workbook.Save(dpiPath);
        Console.WriteLine($"Workbook saved with DPI=300 to {dpiPath}");
        // Restore original DPI
        CellsHelper.DPI = originalDpi;
    }
}