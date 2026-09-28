// Title: Toggle between Normal view and Page Break Preview for an Aspose.Cells worksheet in C#
// AI Prompts: Use reflection to set the ShowPageBreakPreview property on Workbook.Settings according to a boolean variable, ensuring compatibility with older Aspose.Cells versions. | Load an existing .xlsx file or instantiate a new Workbook, change its display mode to either Page Break Preview or Normal, create the output folder if missing, and save the workbook using Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# change worksheet to page break preview programmatically | toggle Excel view between normal and page break preview with Aspose.Cells .NET | reflection workaround for page break preview flag in older Aspose.Cells releases | save workbook after setting view mode using Aspose.Cells C#
// Tags: Aspose.Cells workbook view mode toggle | page break preview setting via reflection | C# load or create Excel workbook Aspose.Cells | ensure output directory exists before saving Aspose.Cells | compatibility handling for page break preview flag Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel file or creates a new workbook, switches its view between Normal and Page Break Preview based on a boolean flag (using reflection for version safety), guarantees the output directory exists, and saves the workbook with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Flag to indicate desired view mode (true = Page Break Preview, false = Normal view)
        bool showPageBreakPreview = true; // Adjust as needed

        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                try
                {
                    workbook = new Workbook(inputPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error loading workbook '{inputPath}': {ex.Message}");
                    return;
                }
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one sheet
                workbook.Worksheets[0].Name = "Sheet1";
            }

            // Attempt to set the view mode using reflection (covers versions without the property)
            try
            {
                var settings = workbook.Settings;
                var prop = settings.GetType().GetProperty("ShowPageBreakPreview");
                if (prop != null && prop.CanWrite)
                {
                    prop.SetValue(settings, showPageBreakPreview);
                }
                else
                {
                    Console.WriteLine("Warning: ShowPageBreakPreview property is not available in this Aspose.Cells version.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Unable to set page break preview mode. {ex.Message}");
            }

            // Ensure the output directory exists
            try
            {
                string? outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not create output directory. {ex.Message}");
            }

            // Save the workbook to the specified output path
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            // General error handling
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
