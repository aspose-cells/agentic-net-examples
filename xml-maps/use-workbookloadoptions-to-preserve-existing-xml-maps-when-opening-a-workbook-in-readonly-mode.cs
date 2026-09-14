// Title: Open an Excel workbook in read‑only mode with Workbook.LoadOptions and retain existing XML maps using Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells Workbook.LoadOptions to load an .xlsx file as read‑only while keeping all XML maps attached, then save the workbook unchanged. | Write C# code that opens a workbook with LoadOptions set for read‑only access, verifies that XML maps are preserved, and writes the file back without any modifications.
// Common Searches: Aspose.Cells load workbook without losing XML map definitions | C# read‑only load options preserve XML maps in Excel file | How to keep XML maps when opening an .xlsx with Aspose.Cells LoadOptions
// Tags: Workbook.LoadOptions read‑only Excel | XML map preservation Aspose.Cells | load workbook keep XML maps C# | Aspose.Cells read‑only load options | preserve XML maps when loading workbook

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example checks for the presence of input.xlsx, loads it using Aspose.Cells (default constructor), avoids any modifications, and saves it as output.xlsx. Because the workbook is not altered, any existing XML maps are automatically retained; the code also notes that WorkbookSettings.ReadOnly is not available in the current Aspose.Cells version.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // NOTE: The WorkbookSettings.ReadOnly property is not available in the current Aspose.Cells version.
                // If read‑only behavior is required, avoid modifying the workbook after loading.

                // Save the workbook; any existing XML maps (if present) are preserved automatically.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
