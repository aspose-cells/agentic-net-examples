// Title: How to disable the Numbers As Text error check in an Aspose.Cells worksheet using C#
// AI Prompts: Write C# code that disables the NumbersAsText validation for a worksheet by setting Workbook.Settings.ErrorCheckOptions.NumbersAsText to false, with a reflection fallback for older Aspose.Cells versions. | Show an example that creates a workbook, turns off the Numbers As Text error check via ErrorCheckOptions, and saves the file to disk. | Provide a C# snippet that detects the presence of the ErrorCheckOptions API and programmatically suppresses the Numbers As Text warning before saving the workbook. | Generate C# code that configures Aspose.Cells workbook settings to ignore the Numbers As Text error and handles missing properties gracefully.
// Common Searches: Aspose.Cells C# disable Numbers As Text error checking on a worksheet | how to turn off numbers as text warning using ErrorCheckOptions in Aspose.Cells | reflection access ErrorCheckOptions NumbersAsText property Aspose.Cells C# | suppress Numbers As Text validation when saving Excel file with Aspose.Cells | disable Excel numbers stored as text error check programmatically C# Aspose
// Tags: disable NumbersAsText error check Aspose.Cells | ErrorCheckOptions configuration C# | reflection set workbook settings property Aspose.Cells | suppress numbers as text warning worksheet | Aspose.Cells workbook save without error validation | C# Excel error checking customization Aspose

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new Workbook, accesses the first worksheet, and uses reflection to locate the Settings.ErrorCheckOptions property. If the property exists, it sets the NumbersAsText flag to false, disabling that specific error check. The code also ensures the output directory exists, saves the workbook as 'output.xlsx', and includes robust exception handling for missing APIs and I/O issues.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (optional, just to demonstrate usage)
            Worksheet sheet = workbook.Worksheets[0];

            // The ErrorCheckOptions API may not be available in all versions.
            // If needed, it can be accessed via reflection; otherwise, we skip it.
            try
            {
                var settingsType = workbook.Settings.GetType();
                var errorCheckProp = settingsType.GetProperty("ErrorCheckOptions");
                if (errorCheckProp != null)
                {
                    var errorCheckOptions = errorCheckProp.GetValue(workbook.Settings);
                    var numbersAsTextProp = errorCheckOptions?.GetType().GetProperty("NumbersAsText");
                    if (numbersAsTextProp != null && numbersAsTextProp.CanWrite)
                    {
                        numbersAsTextProp.SetValue(errorCheckOptions, false);
                    }
                }
            }
            catch (Exception ex)
            {
                // Log or ignore if the property is unavailable in the current library version
                Console.WriteLine("ErrorCheckOptions not supported: " + ex.Message);
            }

            // Define output path
            string outputPath = "output.xlsx";

            // Ensure the directory for the output file exists
            string fullOutputPath = Path.GetFullPath(outputPath);
            string? outputDir = Path.GetDirectoryName(fullOutputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                try
                {
                    Directory.CreateDirectory(outputDir);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to create directory '{outputDir}': {ex.Message}");
                    return;
                }
            }

            // Save the workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception e)
        {
            Console.WriteLine("An unexpected error occurred: " + e.Message);
        }
    }
}
