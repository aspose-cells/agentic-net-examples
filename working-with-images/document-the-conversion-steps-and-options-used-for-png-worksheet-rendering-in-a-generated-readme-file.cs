// Title: C# – Export Excel Worksheet to PNG and Auto‑Generate README with Aspose.Cells
// Description: Creates a workbook, fills sample data, configures ImageOrPrintOptions for high‑resolution PNG (300 DPI, one page per sheet), renders the first sheet page with SheetRender, and writes a README.txt that documents each conversion step and the option values.
// Keywords: Aspose.Cells PNG export C# | ImageOrPrintOptions DPI | SheetRender ToImage example | auto‑generate README from code | Excel to PNG conversion .NET | high‑resolution worksheet image
// Common Searches: export Excel worksheet as PNG using Aspose.Cells C# | set DPI for PNG export with ImageOrPrintOptions | how to create a README that logs Excel image conversion | C# code to render worksheet to PNG with Aspose.Cells | one page per sheet PNG rendering Aspose.Cells
// Developer Intent: Generate a PNG snapshot of a worksheet and produce a README that records the rendering configuration and results.
// Use Cases: Create high‑resolution PNG images of Excel reports for documentation or web publishing. | Maintain an audit trail of export settings (image type, DPI, pagination) by automatically generating a README. | Integrate PNG rendering into CI pipelines where each build logs its conversion parameters for reproducibility.
// AI Prompts: Show a C# script that loops through all worksheets in a workbook and saves each as a 300 DPI PNG using Aspose.Cells. | Provide a markdown README template that lists ImageOrPrintOptions, output file names, and page counts for worksheet‑to‑PNG conversions. | Explain how to modify the example to render every page of a multi‑page worksheet to separate PNG files instead of only the first page.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Creates a workbook, fills sample data, configures ImageOrPrintOptions for high‑resolution PNG (300 DPI, one page per sheet), renders the first sheet page with SheetRender, and writes a README.txt that documents each conversion step and the option values.
class Program
{
    static void Main()
    {
        // Step 1: Create a workbook and add sample data.
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Name");
        worksheet.Cells["B1"].PutValue("Age");
        worksheet.Cells["A2"].PutValue("John");
        worksheet.Cells["B2"].PutValue(30);
        worksheet.Cells["A3"].PutValue("Alice");
        worksheet.Cells["B3"].PutValue(25);

        // Step 2: Configure ImageOrPrintOptions for PNG output.
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = Aspose.Cells.Drawing.ImageType.Png;   // PNG format
        options.OnePagePerSheet = true;                         // Optional: one page per sheet
        options.HorizontalResolution = 300;                     // Optional: 300 DPI
        options.VerticalResolution = 300;                       // Optional: 300 DPI

        // Step 3: Create SheetRender with the worksheet and the options.
        SheetRender sheetRender = new SheetRender(worksheet, options);

        // Step 4: Render the first page of the worksheet to a PNG file.
        string pngFile = "SheetPage0.png";
        sheetRender.ToImage(0, pngFile);

        // Step 5: Build README content that documents the conversion steps and options.
        string readmeContent = "PNG Worksheet Rendering Guide\n";
        readmeContent += "--------------------------------\n";
        readmeContent += "1. Create a Workbook and populate cells with data.\n";
        readmeContent += "2. Create ImageOrPrintOptions and set the following properties:\n";
        readmeContent += "   - ImageType = Png\n";
        readmeContent += "   - OnePagePerSheet = true (forces a single page per sheet)\n";
        readmeContent += "   - HorizontalResolution = 300 DPI (optional)\n";
        readmeContent += "   - VerticalResolution = 300 DPI (optional)\n";
        readmeContent += "3. Instantiate SheetRender using the worksheet and the options.\n";
        readmeContent += "4. Call SheetRender.ToImage(pageIndex, fileName) to generate the PNG.\n";
        readmeContent += $"   Example output file: {pngFile}\n";
        readmeContent += $"   Total pages rendered: {sheetRender.PageCount}\n";

        // Step 6: Write the README file.
        File.WriteAllText("README.txt", readmeContent);

        // Clean up resources.
        sheetRender.Dispose();

        Console.WriteLine("PNG rendering completed and README.txt generated.");
    }
}
