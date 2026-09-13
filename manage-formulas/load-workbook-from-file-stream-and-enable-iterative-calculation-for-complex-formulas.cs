// Title: Load an Excel workbook from a FileStream and enable iterative calculation for complex formulas using Aspose.Cells for .NET
// AI Prompts: Read an Excel file from a FileStream, create a Workbook object, and prepare it for formula processing with Aspose.Cells. | Programmatically turn on iterative calculation, set the iteration count and maximum change values (using reflection if needed) on the workbook's Settings. | Recalculate all formulas after enabling iteration and save the updated workbook to a new file.
// Common Searches: asp.net load excel workbook from filestream aspose.cells | how to enable iterative calculation for circular references in aspose.cells | set iteration count and max change for formula recalculation aspose.cells .net | recalculate all formulas after changing workbook settings aspose.cells | using reflection to modify workbook settings in aspose.cells
// Tags: iterative calculation settings aspose.cells | load workbook from filestream aspose.cells | set iteration count max change aspose.cells | recalculate all formulas aspose.cells | reflection for workbook settings aspose.cells

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates loading an Excel file via a FileStream into an Aspose.Cells Workbook, attempting to enable iterative calculation (including iteration count and max change) using reflection for compatibility, recalculating all formulas, and saving the modified workbook.
class Program
{
    static void Main()
    {
        // Path to the source Excel file
        string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Open a file stream for reading the workbook
            using (FileStream stream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
            {
                // Load the workbook from the stream
                Workbook workbook = new Workbook(stream);

                // Attempt to enable iterative calculation using reflection (covers different API versions)
                try
                {
                    var settings = workbook.Settings;
                    var settingsType = settings.GetType();

                    var iterativeProp = settingsType.GetProperty("IsIterativeCalculationEnabled");
                    if (iterativeProp != null && iterativeProp.CanWrite)
                        iterativeProp.SetValue(settings, true);

                    var iterationCountProp = settingsType.GetProperty("IterationCount");
                    if (iterationCountProp != null && iterationCountProp.CanWrite)
                        iterationCountProp.SetValue(settings, 100);

                    var maxChangeProp = settingsType.GetProperty("MaxChange");
                    if (maxChangeProp != null && maxChangeProp.CanWrite)
                        maxChangeProp.SetValue(settings, 0.001);
                }
                catch (Exception ex)
                {
                    // If reflection fails, continue without iterative settings
                    Console.WriteLine($"Warning: Unable to set iterative calculation settings. {ex.Message}");
                }

                // Recalculate all formulas with the (potential) new settings
                workbook.CalculateFormula();

                // Save the modified workbook (optional)
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook processed and saved to \"{outputPath}\".");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
        }
    }
}
