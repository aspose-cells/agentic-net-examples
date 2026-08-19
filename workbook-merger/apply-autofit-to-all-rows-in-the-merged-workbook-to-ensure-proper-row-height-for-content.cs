// Title: C# – AutoFitRows on Every Sheet of a Merged Workbook Using Aspose.Cells for .NET
// Description: Load a merged Excel file (or create a new workbook if missing), iterate through all worksheets, apply AutoFitRows to adjust row heights to the content, and save the result as a new XLSX file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells AutoFitRows C# | auto adjust row height .NET | fit rows all worksheets | merged workbook row height | Aspose.Cells Excel automation
// Common Searches: auto fit rows in each sheet after merging workbooks Aspose.Cells | C# code to apply AutoFitRows to all worksheets | adjust row height for merged Excel file using Aspose.Cells | Aspose.Cells AutoFitRows example .NET
// Developer Intent: Automatically resize the height of every row in all worksheets of a merged workbook and persist the changes.
// Use Cases: Prepare a consolidated report where wrapped text must be fully visible on every sheet before distribution. | Standardize row heights across multiple combined worksheets for printing or PDF conversion. | Create a final Excel package from several sources and ensure each sheet is optimally formatted without manual adjustments.
// AI Prompts: Generate C# code that loads a workbook, runs AutoFitRows on each worksheet, and saves the file using Aspose.Cells. | Show how to handle a missing source file while applying AutoFitRows to all sheets in a merged workbook with Aspose.Cells for .NET. | Explain the steps to auto‑fit row heights across all worksheets after merging Excel files with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Load a merged Excel file (or create a new workbook if missing), iterate through all worksheets, apply AutoFitRows to adjust row heights to the content, and save the result as a new XLSX file with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            string inputPath = "merged.xlsx";
            string outputPath = "merged_autofit.xlsx";

            Workbook workbook = null;

            try
            {
                // Load existing workbook if it exists; otherwise create a new one
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook(); // creates a workbook with a default sheet
                }

                // AutoFit rows in each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    sheet.AutoFitRows();
                }

                // Save the updated workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
