// Title: C# – Log Count of Worksheets Missing Printer Settings After Batch Deletion with Aspose.Cells
// Description: Creates a workbook, adds three sheets, assigns printer settings to the first sheet, removes "Sheet2", then iterates over the remaining worksheets to count those whose PageSetup.PrinterSettings are null or empty, logs the count, and saves the file as Result.xlsx.
// Keywords: Aspose.Cells C# printer settings | count worksheets without printer settings | log missing printer settings Aspose.Cells | batch worksheet removal .NET | PageSetup.PrinterSettings count | Aspose.Cells worksheet deletion logging
// Common Searches: how to count worksheets without printer settings after deleting a sheet Aspose.Cells | Aspose.Cells log worksheets missing printer settings .NET | remove worksheet and check printer settings count Aspose.Cells | C# Aspose.Cells count sheets with null printer settings | batch delete worksheets and verify printer configuration Aspose
// Developer Intent: Identify and log the number of remaining worksheets that have no printer settings after a batch removal operation.
// Use Cases: Validate printer configuration on all sheets before exporting a workbook. | Generate a quick report of sheets that need manual printer setup. | Ensure workbook integrity after automated worksheet cleanup.
// AI Prompts: Write C# code using Aspose.Cells to delete selected worksheets and return the count of remaining sheets with null or empty PageSetup.PrinterSettings. | Create a reusable method that logs both the number and names of worksheets lacking printer settings after any workbook modification. | Suggest enhancements to capture detailed diagnostics (sheet index, name, timestamp) when logging missing printer settings in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsPrinterSettingsDemo
{
    // Creates a workbook, adds three sheets, assigns printer settings to the first sheet, removes "Sheet2", then iterates over the remaining worksheets to count those whose PageSetup.PrinterSettings are null or empty, logs the count, and saves the file as Result.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Remove the default worksheet to avoid duplicate names
                workbook.Worksheets.Clear();

                // Add worksheets to the workbook
                workbook.Worksheets.Add("Sheet1");
                workbook.Worksheets.Add("Sheet2");
                workbook.Worksheets.Add("Sheet3");

                // Set printer settings for the first worksheet (simulated byte array)
                workbook.Worksheets[0].PageSetup.PrinterSettings = new byte[] { 0x01, 0x02, 0x03 };

                // Batch removal: remove the worksheet named "Sheet2"
                Worksheet sheetToRemove = workbook.Worksheets["Sheet2"];
                if (sheetToRemove != null)
                {
                    // RemoveAt uses the worksheet's index
                    workbook.Worksheets.RemoveAt(sheetToRemove.Index);
                }

                // Count worksheets that lack printer settings after removal
                int worksheetsWithoutPrinterSettings = 0;
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    byte[] printerSettings = ws.PageSetup.PrinterSettings;
                    if (printerSettings == null || printerSettings.Length == 0)
                    {
                        worksheetsWithoutPrinterSettings++;
                    }
                }

                // Log the result
                Console.WriteLine($"Worksheets without printer settings: {worksheetsWithoutPrinterSettings}");

                // Save the workbook
                workbook.Save("Result.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
