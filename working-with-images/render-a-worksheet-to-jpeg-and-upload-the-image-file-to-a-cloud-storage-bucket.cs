using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        try
        {
            // 1. Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");
            sheet.Cells["B1"].PutValue(123);
            sheet.Cells["A2"].PutValue("More Data");
            sheet.Cells["B2"].PutValue(456);

            // 2. Configure image rendering options for JPEG output
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Jpeg,
                OnePagePerSheet = true // render each sheet as a single page
            };

            // 3. Create a SheetRender instance (does not implement IDisposable)
            SheetRender sheetRender = new SheetRender(sheet, imgOptions);

            // 4. Render the first page of the worksheet to a memory stream
            using (MemoryStream imageStream = new MemoryStream())
            {
                sheetRender.ToImage(0, imageStream);
                imageStream.Position = 0; // reset stream position for saving

                // 5. Save the JPEG image locally (replace with cloud upload if needed)
                string outputPath = Path.Combine(Environment.CurrentDirectory, "worksheet_page0.jpg");
                using (FileStream fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                {
                    imageStream.CopyTo(fileStream);
                }

                Console.WriteLine($"Worksheet rendered to JPEG and saved to: {outputPath}");
            }

            // Clean up resources
            workbook.Dispose();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}