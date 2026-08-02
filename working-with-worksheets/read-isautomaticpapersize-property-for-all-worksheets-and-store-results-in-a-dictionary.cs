// Title: Read IsAutomaticPaperSize for All Worksheets into a C# Dictionary with Aspose.Cells
// Description: Shows how to create or load a Workbook, loop through every Worksheet, read the PageSetup.IsAutomaticPaperSize flag, store each result in a Dictionary<string,bool> keyed by worksheet name, display the values, and optionally save the workbook as XLSX using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | IsAutomaticPaperSize | C# | .NET | worksheet page setup | read property | dictionary | iterate worksheets | print values | save workbook
// Common Searches: Aspose.Cells read IsAutomaticPaperSize property | C# get automatic paper size for each worksheet | store worksheet page setup flags in dictionary | iterate worksheets Aspose.Cells .NET | how to check IsAutomaticPaperSize in Aspose.Cells
// Developer Intent: Extract the IsAutomaticPaperSize flag from every worksheet and map it to the worksheet name.
// Use Cases: Create a summary report of which sheets use automatic paper sizing before printing. | Apply conditional page‑setup changes based on each sheet’s automatic paper size setting. | Log or export paper‑size configurations for compliance or documentation purposes.
// AI Prompts: Generate C# code using Aspose.Cells that iterates through all worksheets, reads PageSetup.IsAutomaticPaperSize, and returns a Dictionary<string,bool>. | Provide an example that captures IsAutomaticPaperSize for each sheet, prints the results, and saves the workbook. | Explain how to toggle the IsAutomaticPaperSize property for selected worksheets based on a custom condition.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to create or load a Workbook, loop through every Worksheet, read the PageSetup.IsAutomaticPaperSize flag, store each result in a Dictionary<string,bool> keyed by worksheet name, display the values, and optionally save the workbook as XLSX using Aspose.Cells for .NET.
    public class ReadIsAutomaticPaperSizeDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (or load an existing file if needed)
                Workbook workbook = new Workbook();

                // Add a second worksheet for demonstration
                workbook.Worksheets.Add();

                // Dictionary to hold worksheet name and its IsAutomaticPaperSize value
                Dictionary<string, bool> automaticPaperSizeMap = new Dictionary<string, bool>();

                // Iterate through all worksheets in the workbook
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    Worksheet sheet = workbook.Worksheets[i];
                    // Read the IsAutomaticPaperSize property from the worksheet's PageSetup
                    bool isAutomatic = sheet.PageSetup.IsAutomaticPaperSize;
                    automaticPaperSizeMap[sheet.Name] = isAutomatic;
                }

                // Output the collected values
                foreach (var kvp in automaticPaperSizeMap)
                {
                    Console.WriteLine($"Worksheet \"{kvp.Key}\" - Automatic Paper Size: {kvp.Value}");
                }

                // Save the workbook (optional, demonstrates lifecycle usage)
                string outputPath = "ReadIsAutomaticPaperSizeDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to \"{Path.GetFullPath(outputPath)}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ReadIsAutomaticPaperSizeDemo.Run();
        }
    }
}
