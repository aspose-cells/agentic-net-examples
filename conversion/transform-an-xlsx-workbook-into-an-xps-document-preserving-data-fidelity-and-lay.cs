using System;
using Aspose.Cells;
using Aspose.Cells.Rendering; // Required for XpsSaveOptions inheritance
using Aspose.Cells.Saving;   // Namespace for SaveOptions base class (if needed)

class XlsxToXpsConverter
{
    static void Main()
    {
        // Path to the source XLSX workbook
        string sourcePath = "input.xlsx";

        // Path for the resulting XPS document
        string destPath = "output.xps";

        // Load the workbook from the XLSX file
        Workbook workbook = new Workbook(sourcePath);

        // Create XPS save options to control the conversion
        XpsSaveOptions xpsOptions = new XpsSaveOptions
        {
            // Preserve the original layout by rendering each sheet on a separate page
            OnePagePerSheet = false,

            // Ensure all columns of a sheet are kept on a single page if desired
            // AllColumnsInOnePagePerSheet = true,

            // Use a default font that supports Unicode characters
            DefaultFont = "Arial",

            // Enable font compatibility checks for accurate rendering
            CheckFontCompatibility = true,

            // Render the entire workbook (no page range restriction)
            PageIndex = 0,
            PageCount = int.MaxValue
        };

        // Save the workbook as XPS using the specified options
        workbook.Save(destPath, xpsOptions);

        Console.WriteLine($"Conversion completed: '{sourcePath}' -> '{destPath}'");
    }
}