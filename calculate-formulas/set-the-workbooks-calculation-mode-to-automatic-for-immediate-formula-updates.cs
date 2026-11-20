// Title: How to set Aspose.Cells workbook calculation mode to Automatic in C# for instant formula updates
// Common Searches: Aspose.Cells C# set workbook calculation mode to Automatic | Enable automatic formula recalculation in Aspose.Cells .NET | Change CalcMode to Automatic for an Aspose.Cells workbook | Make Excel formulas recalculate instantly using Aspose.Cells C# | Configure workbook settings for automatic calculation in Aspose.Cells
// Tags: Aspose.Cells set CalcMode to Automatic | C# workbook calculation mode configuration | automatic formula evaluation Aspose.Cells | Excel workbook automatic recalculation .NET | Aspose.Cells workbook settings API

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new Workbook, shows how to explicitly set workbook.Settings.CalcMode to CalcMode.Automatic so formulas recalculate immediately, ensures the output directory exists, and saves the file as Result.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // The default calculation mode is Automatic, so no explicit setting is required.
                // If you need to change it in the future, use:
                // workbook.Settings.CalcMode = CalcMode.Automatic;

                // Define the output file path
                string outputPath = "Result.xlsx";

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
