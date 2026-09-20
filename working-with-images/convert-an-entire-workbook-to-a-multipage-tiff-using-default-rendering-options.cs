// Title: How to convert an entire Excel workbook to a multi‑page TIFF using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx workbook and saves it as a multi‑page TIFF where each worksheet becomes a separate page, using Aspose.Cells default settings. | Provide a step‑by‑step example of exporting all worksheets of an Excel file to a single TIFF image with Aspose.Cells, without customizing rendering options.
// Common Searches: Aspose.Cells C# save whole workbook as multi page TIFF file | Convert Excel workbook to multi page TIFF with default options in .NET | How to generate a TIFF file where each Excel sheet is a separate page using Aspose.Cells | Example code for saving .xlsx as multi‑page TIFF with Aspose.Cells | Default rendering of TIFF images from Excel workbooks in Aspose.Cells
// Tags: Aspose.Cells workbook to multi-page TIFF conversion | C# save Excel as TIFF using Aspose.Cells | default TIFF rendering options Aspose.Cells | export worksheets to separate TIFF pages .NET | multi-page TIFF generation from Excel workbook

using Aspose.Cells;

// The sample loads an Excel workbook (input.xlsx) with Aspose.Cells and saves the entire workbook as a multi‑page TIFF (output.tiff) using default rendering options, rendering each worksheet as an individual page in the TIFF file.
class Program
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Save the entire workbook as a multi‑page TIFF using default rendering options
        // Each worksheet will be rendered as a separate page in the TIFF file
        workbook.Save("output.tiff", SaveFormat.Tiff);
    }
}
