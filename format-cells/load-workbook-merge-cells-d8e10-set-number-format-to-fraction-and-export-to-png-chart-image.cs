// Title: C# – Merge D8:E10, Apply Fraction Format, and Export Worksheet to PNG with Aspose.Cells
// Description: Load or create an Excel workbook, merge the range D8:E10 on the first sheet, set a custom fraction number format ("# ?/??"), insert a sample value, render the first page as a PNG image, and optionally save the workbook. All operations are performed using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# merge cells | Excel D8:E10 merge | fraction number format Aspose | export worksheet to PNG | WorkbookRender PNG Aspose.Cells | custom number format # ?/?? | C# Excel to image conversion | Aspose.Cells render page as image | save workbook as PNG | Aspose.Cells example
// Common Searches: How to merge a cell range and set a fraction format with Aspose.Cells .NET | Export an Excel sheet to PNG after merging cells using C# | Apply custom number format '# ?/??' to merged cells in Aspose.Cells | Render first worksheet page to PNG image in C# | Aspose.Cells example for merging cells and image export
// Developer Intent: Merge cells D8:E10, format them as a fraction, and generate a PNG snapshot of the worksheet using Aspose.Cells for .NET.
// Use Cases: Create printable reports where merged cells display fractional values and are delivered as PNG thumbnails for web portals. | Automate conversion of Excel dashboards with merged cells into PNG images for email newsletters or documentation. | Generate visual previews of formatted Excel data for UI components or API responses.
// AI Prompts: Provide C# code that uses Aspose.Cells to merge D8:E10, apply the fraction format "# ?/??", insert a numeric value, and render the first worksheet page to a PNG file. | Show an Aspose.Cells for .NET example that loads an existing workbook, merges a cell range, sets a custom fraction number format, saves the workbook, and exports the sheet as a PNG image.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Load or create an Excel workbook, merge the range D8:E10 on the first sheet, set a custom fraction number format ("# ?/??"), insert a sample value, render the first page as a PNG image, and optionally save the workbook. All operations are performed using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load an existing workbook if it exists; otherwise create a new one
        string inputPath = "input.xlsx";
        Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

        // Access the first worksheet and its cells collection
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Merge cells D8:E10
        // D8 -> row 7, column 3 (zero‑based)
        // Span 3 rows (8,9,10) and 2 columns (D,E)
        cells.Merge(7, 3, 3, 2);

        // Set the number format of the merged cell to a fraction
        Style mergedStyle = cells[7, 3].GetStyle();
        mergedStyle.Custom = "# ?/??";          // Fraction format
        cells[7, 3].SetStyle(mergedStyle);

        // Put a sample numeric value to demonstrate the fraction format
        cells[7, 3].PutValue(0.75);

        // Render the worksheet (first page) to a PNG image
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
        {
            ImageType = Aspose.Cells.Drawing.ImageType.Png,
            OnePagePerSheet = true
        };
        WorkbookRender renderer = new WorkbookRender(workbook, imgOptions);
        string imagePath = "output.png";
        renderer.ToImage(0, imagePath);   // Render page 0 to the PNG file

        // Save the workbook (optional)
        workbook.Save("output.xlsx");
    }
}
