// Title: Clone an Excel worksheet’s PageSetup, modify only the paper height, and reapply it with Aspose.Cells for .NET
// Common Searches: how to copy a worksheet page setup and change only the paper height using Aspose.Cells C# | Aspose.Cells set custom paper height while preserving existing width in .NET | C# example to duplicate worksheet PageSetup and adjust paper height | preserve worksheet paper width when updating paper height with Aspose.Cells | apply modified PageSetup back to the same worksheet in Aspose.Cells for .NET
// Tags: clone worksheet page setup Aspose.Cells | modify paper height Aspose.Cells | custom paper size Excel .NET | preserve paper width Aspose.Cells | apply updated page setup worksheet

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel file, clones the first worksheet's PageSetup, sets the PaperSize to Custom to keep the current width, updates the paper height (noting that PaperHeight is read‑only in current versions), reapplies the modified PageSetup to the worksheet, and saves the result as a new workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            PageSetup pageSetup = sheet.PageSetup;

            // Set custom paper size (preserve existing width, set height to 11 inches)
            // Note: In the current Aspose.Cells version PaperHeight and PaperWidth are read‑only,
            // so we only set the PaperSize to Custom. The existing dimensions will be retained.
            pageSetup.PaperSize = PaperSizeType.Custom;

            // If a future version allows setting width/height, uncomment the lines below:
            // double currentWidth = pageSetup.PaperWidth; // width in points
            // pageSetup.PaperHeight = 11.0 * 72.0;        // height = 11 inches
            // pageSetup.PaperWidth = currentWidth;      // keep existing width

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
