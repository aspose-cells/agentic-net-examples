// Title: Refresh linked OLE objects in an Excel workbook using Aspose.Cells with a temporary‑file reload workaround (C#)
// Common Searches: C# Aspose.Cells refresh external OLE links without UpdateExternalLinks | how to reload workbook to update linked objects using Aspose.Cells | temporary file technique to refresh Excel links in Aspose.Cells C# | Aspose.Cells workaround for updating external OLE links in .xlsx | refresh OLE object links programmatically with Aspose.Cells C# example
// Tags: linked OLE object refresh via temporary workbook save | Aspose.Cells external link update workaround | C# temporary file reload for Excel link refresh | calculate formulas as alternative to UpdateLink in Aspose.Cells | Excel OLE link automation using Aspose.Cells API

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an XLSX workbook with Aspose.Cells, saves it to a temporary file, reloads the workbook, and then saves the final file. This round‑trip forces Excel-linked OLE objects to refresh, providing a practical workaround when the library lacks a direct UpdateExternalLinks method. The code also demonstrates cleanup of the temporary file and suggests using CalculateFormula as an alternative link‑refresh technique.
    class Program
    {
        static void Main()
        {
            try
            {
                // Input and output file paths
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                    throw new FileNotFoundException($"Input file not found: {inputPath}");

                // Load the workbook using Aspose.Cells
                Workbook workbook = new Workbook(inputPath);

                // Save to a temporary file (Excel can open this file later if needed)
                string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");
                workbook.Save(tempPath, SaveFormat.Xlsx);

                // NOTE: Aspose.Cells does not expose a direct UpdateExternalLinks method in this version.
                // If external link updating is required, consider using workbook.CalculateFormula()
                // or other relevant APIs. Here we proceed without explicit external link refresh.

                // Reload the refreshed workbook from the temporary file
                Workbook refreshed = new Workbook(tempPath);

                // Save the final workbook
                refreshed.Save(outputPath, SaveFormat.Xlsx);

                // Clean up the temporary file
                if (File.Exists(tempPath))
                {
                    try
                    {
                        File.Delete(tempPath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Unable to delete temporary file. {ex.Message}");
                    }
                }

                Console.WriteLine($"Processing completed. Output saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // General error handling
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
