using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    class WorkbookToPng
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // Step 1: Create a sample workbook and save it into a memory stream
            // ------------------------------------------------------------
            Workbook sourceWorkbook = new Workbook();
            Worksheet ws = sourceWorkbook.Worksheets[0];
            ws.Cells["A1"].PutValue("Hello Aspose.Cells");
            ws.Cells["B2"].PutValue(12345);

            // Save the workbook to a memory stream (XLSX format)
            using (MemoryStream sourceStream = new MemoryStream())
            {
                sourceWorkbook.Save(sourceStream, SaveFormat.Xlsx);

                // Reset the stream position before reading
                sourceStream.Position = 0;

                // ------------------------------------------------------------
                // Step 2: Load the workbook from the memory stream
                // ------------------------------------------------------------
                Workbook loadedWorkbook = new Workbook(sourceStream);

                // ------------------------------------------------------------
                // Step 3: Configure image rendering options (PNG, 300 DPI)
                // ------------------------------------------------------------
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Png,          // Output format
                    HorizontalResolution = 300,         // 300 DPI horizontally
                    VerticalResolution = 300,           // 300 DPI vertically
                    OnePagePerSheet = true              // Render each sheet as a single page
                };

                // ------------------------------------------------------------
                // Step 4: Render the whole workbook to a PNG image file
                // ------------------------------------------------------------
                // The WorkbookRender.ToImage(string) method renders the entire workbook.
                // It respects the ImageOrPrintOptions set above.
                WorkbookRender renderer = new WorkbookRender(loadedWorkbook, imgOptions);
                string outputPath = "WorkbookImage.png";
                renderer.ToImage(outputPath);

                Console.WriteLine($"Workbook rendered to PNG with 300 DPI at: {outputPath}");
            }
        }
    }
}