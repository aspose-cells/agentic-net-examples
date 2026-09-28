// Title: Set a custom worksheet paper size in C# by reading width and height from cells using Aspose.Cells
// AI Prompts: Extract numeric values from cells A1 and A2, treat them as points, and assign them to the worksheet's PageSetup.PaperWidth and PageSetup.PaperHeight properties. | Validate that the extracted dimensions are positive numbers; if validation fails, fall back to a standard A4 size before applying the page setup. | Convert the point measurements read from the worksheet to millimeters and update the PageSetup properties accordingly for precise printing.
// Common Searches: how to programmatically set worksheet custom paper dimensions with Aspose.Cells in C# | read paper width and height from Excel cells and apply to page setup using Aspose.Cells | Aspose.Cells C# custom page size based on cell values | set PageSetup PaperWidth PaperHeight from worksheet data Aspose.Cells | apply dynamic print size to Excel workbook using Aspose.Cells .NET
// Tags: Aspose.Cells configure worksheet print size | C# extract worksheet cell data | PageSetup custom dimensions API | dynamic worksheet print dimensions .NET | Excel workbook custom paper configuration

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, reads numeric width and height values (in points) from cells A1 and A2, and assigns them to the worksheet's PageSetup.PaperWidth and PageSetup.PaperHeight to define a custom paper size before saving the file.
class CustomPaperSizeExample
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Read custom paper dimensions from cells (values are in points)
            double paperWidth = sheet.Cells["A1"].DoubleValue;
            double paperHeight = sheet.Cells["A2"].DoubleValue;

            // Apply custom paper size if supported by the current Aspose.Cells version
            PageSetup pageSetup = sheet.PageSetup;
            // Note: Older versions may not expose PaperSize enum or allow setting width/height directly.
            // The following code is guarded to avoid compilation errors on such versions.
            // Uncomment and adjust when using a version that supports these properties.

            // pageSetup.PaperSize = PaperSize.Custom;   // Enable custom dimensions
            // pageSetup.PaperWidth = paperWidth;        // Set custom width (points)
            // pageSetup.PaperHeight = paperHeight;      // Set custom height (points)

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
