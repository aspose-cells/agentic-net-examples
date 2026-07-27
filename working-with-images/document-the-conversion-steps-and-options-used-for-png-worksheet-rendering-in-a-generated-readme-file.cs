using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsReadmeGenerator
{
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // Step 1: Create a new workbook and add sample data.
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();                     // Create a workbook (provided constructor)
            Worksheet sheet = workbook.Worksheets[0];              // Get the first worksheet

            // Populate some cells with data
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Alice");
            sheet.Cells["B3"].PutValue(25);

            // -----------------------------------------------------------------
            // Step 2: Configure ImageOrPrintOptions for PNG rendering.
            // -----------------------------------------------------------------
            ImageOrPrintOptions options = new ImageOrPrintOptions(); // Provided constructor
            options.ImageType = ImageType.Png;                       // Set output image type to PNG
            options.OnePagePerSheet = true;                         // Render each sheet as a single page
            options.HorizontalResolution = 300;                     // Optional: set DPI for higher quality
            options.VerticalResolution = 300;                       // Optional: set DPI for higher quality

            // -----------------------------------------------------------------
            // Step 3: Create SheetRender using the worksheet and options.
            // -----------------------------------------------------------------
            SheetRender renderer = new SheetRender(sheet, options); // Provided constructor

            // Retrieve page count (useful information for the README)
            int pageCount = renderer.PageCount;

            // Render the first page to a PNG file (demonstrates the rendering process)
            string imagePath = "RenderedSheet.png";
            renderer.ToImage(0, imagePath);                         // Provided method overload

            // -----------------------------------------------------------------
            // Step 4: Build README content describing the conversion steps.
            // -----------------------------------------------------------------
            string readmeContent =
$@"# PNG Worksheet Rendering with Aspose.Cells

This README documents the steps and options used to convert an Excel worksheet to a PNG image.

## Steps Performed

1. **Create Workbook**
   - `Workbook workbook = new Workbook();`
   - Added sample data to cells A1:B3.

2. **Configure ImageOrPrintOptions**
   - `options.ImageType = ImageType.Png;`   // Output format
   - `options.OnePagePerSheet = true;`     // Render whole sheet on one page
   - `options.HorizontalResolution = 300;` // DPI (optional)
   - `options.VerticalResolution = 300;`   // DPI (optional)

3. **Create SheetRender**
   - `SheetRender renderer = new SheetRender(sheet, options);`

4. **Render to PNG**
   - `renderer.ToImage(0, ""{imagePath}"");`
   - The workbook contains **{pageCount}** page(s); page index `0` was rendered.

## Result

- PNG image saved at: `{Path.GetFullPath(imagePath)}`
- README file generated at: `{Path.GetFullPath("README.txt")}`

Feel free to modify the `ImageOrPrintOptions` properties to suit your rendering requirements (e.g., change DPI, enable/disable `OnePagePerSheet`, etc.).";

            // -----------------------------------------------------------------
            // Step 5: Write the README file to disk.
            // -----------------------------------------------------------------
            string readmePath = "README.txt";
            File.WriteAllText(readmePath, readmeContent); // Standard .NET file write

            // Clean up resources
            renderer.Dispose();

            Console.WriteLine("PNG rendering completed and README generated.");
        }
    }
}