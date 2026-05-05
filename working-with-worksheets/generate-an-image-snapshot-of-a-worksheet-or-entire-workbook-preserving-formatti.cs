using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

public class WorkbookSnapshotDemo
{
    public static void Run()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Name = "SampleSheet";
        worksheet.Cells["A1"].PutValue("Header");
        worksheet.Cells["A2"].PutValue("Item 1");
        worksheet.Cells["B2"].PutValue(100);
        worksheet.Cells["A3"].PutValue("Item 2");
        worksheet.Cells["B3"].PutValue(200);

        // Configure image rendering options
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            // Default image format is PNG, no need to set explicitly
            OnePagePerSheet = true // Render each sheet on a separate page
        };

        // Ensure output directory exists
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Render the entire workbook to an image file
        WorkbookRender workbookRenderer = new WorkbookRender(workbook, options);
        string workbookImagePath = Path.Combine(outputDir, "workbook_snapshot.png");
        using (FileStream fs = new FileStream(workbookImagePath, FileMode.Create, FileAccess.Write))
        {
            workbookRenderer.ToImage(0, fs);
        }
        Console.WriteLine($"Workbook rendered to image: {workbookImagePath}");

        // Render the first worksheet (single page) to an image file
        SheetRender sheetRenderer = new SheetRender(worksheet, options);
        string worksheetImagePath = Path.Combine(outputDir, "worksheet_snapshot.png");
        using (FileStream fs = new FileStream(worksheetImagePath, FileMode.Create, FileAccess.Write))
        {
            sheetRenderer.ToImage(0, fs);
        }
        Console.WriteLine($"Worksheet rendered to image: {worksheetImagePath}");
    }
}

public class Program
{
    public static void Main()
    {
        WorkbookSnapshotDemo.Run();
    }
}