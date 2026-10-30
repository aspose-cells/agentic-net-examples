// Title: Export an Excel worksheet to a 150 DPI TIFF image using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# save worksheet as TIFF with 150 DPI resolution | How to set horizontal and vertical DPI when converting Excel to TIFF in .NET | Render Excel sheet to high‑resolution TIFF using ImageOrPrintOptions | C# code to export multiple worksheets to separate 150 DPI TIFF files with Aspose.Cells | OnePagePerSheet option for TIFF export in Aspose.Cells
// Tags: Aspose.Cells ImageOrPrintOptions TIFF DPI | C# export Excel to high‑resolution TIFF | SheetRender render worksheet to TIFF | OnePagePerSheet TIFF rendering Aspose.Cells | multiple worksheets TIFF export Aspose.Cells

// Load an existing workbook
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook("input.xlsx");

// Configure image rendering options
Aspose.Cells.Rendering.ImageOrPrintOptions imgOptions = new Aspose.Cells.Rendering.ImageOrPrintOptions
{
    // Set the output format to TIFF
    SaveFormat = Aspose.Cells.SaveFormat.Tiff,
    // Set horizontal and vertical DPI to 150
    HorizontalResolution = 150,
    VerticalResolution = 150,
    // Optional: render the whole sheet on a single page
    OnePagePerSheet = true
};

// Render the first worksheet to a TIFF image
Aspose.Cells.Rendering.SheetRender sheetRender = new Aspose.Cells.Rendering.SheetRender(workbook.Worksheets[0], imgOptions);
sheetRender.ToImage(0, "output.tiff");

// If you need to render all worksheets, iterate through them:
// for (int i = 0; i < workbook.Worksheets.Count; i++)
// {
//     Aspose.Cells.Rendering.SheetRender sr = new Aspose.Cells.Rendering.SheetRender(workbook.Worksheets[i], imgOptions);
//     sr.ToImage(0, $"output_sheet{i + 1}.tiff");
// }
