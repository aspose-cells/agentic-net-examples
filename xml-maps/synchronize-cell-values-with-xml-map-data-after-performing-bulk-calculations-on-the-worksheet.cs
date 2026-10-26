// Title: Recalculate all worksheet formulas and prepare workbook for XML map synchronization using Aspose.Cells for .NET (C#)
// Common Searches: how to recalculate all formulas in an Excel workbook with Aspose.Cells C# | updating XML map values after bulk calculations in Aspose.Cells .NET | saving workbook after formula recalculation before exporting XML map data using Aspose.Cells | Aspose.Cells C# example for bulk worksheet calculations and XML map preparation | workaround for missing XmlMaps API in Aspose.Cells when syncing XML data
// Tags: recalculate formulas Aspose.Cells .NET | XML map synchronization Aspose.Cells | bulk worksheet calculation C# | export XML map data Aspose.Cells | missing XmlMaps API workaround

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The program loads an existing workbook (or creates a new one), recalculates all formulas, notes that the current Aspose.Cells version lacks XmlMaps/XmlMap.ExportData support, and saves the workbook, ready for XML map synchronization.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "InputWorkbook.xlsx";
            string outputPath = "OutputWorkbook.xlsx";

            try
            {
                // Load existing workbook or create a new blank one if the file is missing.
                Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

                // Recalculate all formulas in the workbook.
                workbook.CalculateFormula();

                // Get the first worksheet (adjust index if needed).
                Worksheet sheet = workbook.Worksheets[0];

                // -----------------------------------------------------------------
                // XML map handling – removed because the current Aspose.Cells version
                // does not expose Workbook.XmlMaps / XmlMap.ExportData APIs.
                // If a future version adds this support, the code can be re‑added
                // using the appropriate API calls.
                // -----------------------------------------------------------------

                // Ensure the directory for the output file exists.
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the updated workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
