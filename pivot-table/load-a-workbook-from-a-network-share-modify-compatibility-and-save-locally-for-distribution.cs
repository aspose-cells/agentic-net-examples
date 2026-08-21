// Title: Load Excel from UNC share, set compatibility, and save locally using Aspose.Cells (C#)
// Description: C# example that checks a workbook's existence on a UNC network share, loads it with Aspose.Cells, enables the CheckCompatibility flag for older Excel versions, creates the target folder if needed, and saves the file as XLSX to a local path while handling errors.
// Keywords: Aspose.Cells UNC path | load workbook from network share C# | CheckCompatibility Aspose.Cells | save Excel locally Aspose.Cells | C# Excel file network share example | Aspose.Cells file existence check | Create directory before saving Aspose.Cells
// Common Searches: How to open an Excel file from a UNC path with Aspose.Cells | Aspose.Cells enable compatibility check before saving | Save Aspose.Cells workbook to a specific local folder | C# verify network file exists before loading Excel | Create missing directory when saving Aspose.Cells workbook
// Developer Intent: Open an Excel workbook located on a network share, turn on compatibility checking, and write the modified file to a local directory.
// Use Cases: Distribute a company‑wide template stored on a shared drive, ensuring it remains compatible with Excel 97‑2003 before sending copies to users. | Automate nightly processing that reads workbooks from a file server, applies compatibility settings, and stages them in a local folder for downstream workflows. | Prevent runtime errors in a server application by confirming the network file exists and the output folder is present before using Aspose.Cells.
// AI Prompts: Generate C# code that loads an Excel workbook from a UNC path with Aspose.Cells, sets workbook.Settings.CheckCompatibility = true, creates the destination folder if missing, and saves the file locally. | Show how to handle FileNotFound and other exceptions when opening a workbook from a network share using Aspose.Cells. | Explain the purpose of the CheckCompatibility property in Aspose.Cells and when to apply it before saving a workbook.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsNetworkExample
{
    // C# example that checks a workbook's existence on a UNC network share, loads it with Aspose.Cells, enables the CheckCompatibility flag for older Excel versions, creates the target folder if needed, and saves the file as XLSX to a local path while handling errors.
    class Program
    {
        static void Main()
        {
            // Path to the workbook on a network share
            string networkFilePath = @"\\ServerName\ShareFolder\SourceWorkbook.xlsx";

            // Verify that the source file exists before attempting to load it
            if (!File.Exists(networkFilePath))
            {
                Console.WriteLine($"Source workbook not found at: {networkFilePath}");
                return;
            }

            try
            {
                // Load the workbook from the network location
                Workbook workbook = new Workbook(networkFilePath);

                // Enable compatibility checks for older Excel versions
                workbook.Settings.CheckCompatibility = true;

                // Define local path for the modified workbook
                string localSavePath = @"C:\Temp\ModifiedWorkbook.xlsx";

                // Ensure the target directory exists
                string localDir = Path.GetDirectoryName(localSavePath);
                if (!Directory.Exists(localDir))
                {
                    Directory.CreateDirectory(localDir);
                }

                // Save the modified workbook locally
                workbook.Save(localSavePath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook loaded, compatibility modified, and saved to: {localSavePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
