// Title: Convert an Excel worksheet to a PNG image with visible gridlines using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# render worksheet to PNG while keeping gridlines visible | How to export Excel sheet as PNG image with gridlines using Aspose.Cells .NET | C# code to save Excel worksheet as PNG preserving gridlines | Generate PNG from Excel workbook with gridlines using Aspose.Cells | Render Excel worksheet to PNG with one page per sheet Aspose.Cells
// Tags: worksheet to PNG rendering Aspose.Cells | preserve gridlines image export | ImageOrPrintOptions PNG OnePagePerSheet | SheetRender C# Excel to image | IsGridlinesVisible property Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;   // For ImageType enum

// The example loads an Excel file (or creates a sample workbook), enables worksheet gridlines via IsGridlinesVisible, configures ImageOrPrintOptions for PNG output with OnePagePerSheet, and uses SheetRender to generate PNG images that retain the original gridline appearance.
class Program
{
    static void Main(string[] args)
    {
        try
        {
            string inputPath;
            string outputPath;

            if (args.Length >= 2)
            {
                inputPath = args[0];
                outputPath = args[1];

                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                    return;
                }
            }
            else
            {
                // No arguments – create a sample workbook in memory.
                Console.WriteLine("No input arguments supplied. Creating a sample workbook.");
                Workbook sampleWb = CreateSampleWorkbook();

                inputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");
                sampleWb.Save(inputPath, SaveFormat.Xlsx);
                outputPath = "sample_output.png";

                Console.WriteLine($"Sample workbook saved to temporary file: {inputPath}");
            }

            // Load the workbook.
            Workbook wb = new Workbook(inputPath);
            Worksheet sheet = wb.Worksheets[0]; // Convert the first worksheet.

            // Ensure gridlines are visible in the worksheet.
            sheet.IsGridlinesVisible = true;

            // Configure image rendering options.
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png,   // Output format PNG.
                OnePagePerSheet = true,      // Render the whole sheet on a single page.
                Transparent = false          // Opaque background (optional).
                // GridLines option is not required when worksheet gridlines are visible.
            };

            // Create the renderer.
            SheetRender renderer = new SheetRender(sheet, imgOptions);

            // Render each page. Usually there will be only one page because of OnePagePerSheet=true.
            for (int page = 0; page < renderer.PageCount; page++)
            {
                try
                {
                    // Build a filename for each page if there are multiple pages.
                    string pageOutput = renderer.PageCount == 1
                        ? outputPath
                        : Path.Combine(
                            Path.GetDirectoryName(outputPath) ?? "",
                            Path.GetFileNameWithoutExtension(outputPath) + $"_page{page + 1}" + Path.GetExtension(outputPath));

                    renderer.ToImage(page, pageOutput);
                    Console.WriteLine($"Page {page + 1} rendered to \"{pageOutput}\"");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error rendering page {page + 1}: {ex.Message}");
                }
            }

            Console.WriteLine("Conversion completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }

    private static Workbook CreateSampleWorkbook()
    {
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Populate sample data.
        ws.Cells["A1"].PutValue("Product");
        ws.Cells["B1"].PutValue("Quantity");
        ws.Cells["C1"].PutValue("Price");

        ws.Cells["A2"].PutValue("Apple");
        ws.Cells["B2"].PutValue(120);
        ws.Cells["C2"].PutValue(0.5);

        ws.Cells["A3"].PutValue("Banana");
        ws.Cells["B3"].PutValue(85);
        ws.Cells["C3"].PutValue(0.3);

        ws.Cells["A4"].PutValue("Cherry");
        ws.Cells["B4"].PutValue(60);
        ws.Cells["C4"].PutValue(1.2);

        // Apply a simple style to make the gridlines more noticeable.
        Style headerStyle = wb.CreateStyle();
        headerStyle.Font.IsBold = true;
        headerStyle.ForegroundColor = System.Drawing.Color.LightGray;
        headerStyle.Pattern = BackgroundType.Solid;

        ws.Cells.CreateRange("A1:C1").ApplyStyle(headerStyle, new StyleFlag { FontBold = true, CellShading = true });

        // Ensure gridlines are visible.
        ws.IsGridlinesVisible = true;

        return wb;
    }
}
