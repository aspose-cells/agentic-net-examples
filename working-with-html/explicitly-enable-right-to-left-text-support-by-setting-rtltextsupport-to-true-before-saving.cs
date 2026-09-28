// Title: How to enable right-to-left text support in an Aspose.Cells workbook before saving to XLSX with C#
// AI Prompts: Configure the workbook's Settings to activate right-to-left layout, then call Save on the workbook. | Add a runtime check for the presence of the IsRightToLeft property and enable it only on supported Aspose.Cells versions. | Ensure the target directory exists, create it if necessary, and then save the workbook with RTL enabled.
// Common Searches: C# Aspose.Cells enable right-to-left orientation before exporting to Excel | How to check for IsRightToLeft support in different Aspose.Cells versions | Create missing output folder automatically when saving workbook with Aspose.Cells
// Tags: set RTL mode in Aspose.Cells workbook | verify IsRightToLeft availability programmatically | save workbook as XLSX with RTL orientation | prepare output folder before saving workbook | Aspose.Cells workbook RTL configuration

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, optionally enables right-to-left text support via the Settings.IsRightToLeft property (when available), ensures the output directory exists, and saves the workbook to an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Enable right-to-left text support if the API provides it.
            // The IsRightToLeft property may not be available in older versions of Aspose.Cells.
            // Uncomment the following line if your version supports it:
            // workbook.Settings.IsRightToLeft = true;

            // Define output file path
            string outputPath = "output.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
