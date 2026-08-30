// Title: Load an Excel workbook from a UNC network share, set CheckCompatibility, and save it locally using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an .xlsx file from a UNC path, turns on workbook.Settings.CheckCompatibility, and writes the file to a local folder with Aspose.Cells. | Create a C# console program that validates the existence of a workbook on a network share, creates the destination directory if missing, and saves the workbook as XLSX after enabling compatibility mode using Aspose.Cells. | Show how to use Aspose.Cells to read a workbook from a network location, activate Excel compatibility checks, and export the modified workbook to a specified local path.
// Common Searches: aspnet load excel file from UNC path using Aspose.Cells | how to enable CheckCompatibility in Aspose.Cells before saving workbook | save workbook to local folder after opening from network share with Aspose.Cells C# | verify network file exists before loading Aspose.Cells | create destination directory automatically when saving workbook Aspose.Cells
// Tags: read Excel file from network share Aspose.Cells | set workbook check compatibility Aspose.Cells | save workbook as XLSX locally Aspose.Cells | verify network file existence C# Aspose.Cells | create destination directory before save Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsNetworkExample
{
    // // Loads an Excel workbook from a UNC network share, enables compatibility checks, ensures the target folder exists, and saves the workbook locally as XLSX using Aspose.Cells.
    class Program
    {
        static void Main()
        {
            // Path to the workbook on a network share (UNC path)
            string networkPath = @"\\Server\Share\Documents\SourceWorkbook.xlsx";

            // Local path where the modified workbook will be saved
            string localPath = @"C:\Temp\ModifiedWorkbook.xlsx";

            // Verify that the source workbook exists before attempting to load it
            if (!File.Exists(networkPath))
            {
                Console.WriteLine($"Source workbook not found at network location: {networkPath}");
                return;
            }

            try
            {
                // Load the workbook from the network location
                using (Workbook workbook = new Workbook(networkPath))
                {
                    // Enable compatibility checks for older Excel versions
                    workbook.Settings.CheckCompatibility = true;

                    // Ensure the target directory exists
                    string localDir = Path.GetDirectoryName(localPath);
                    if (!Directory.Exists(localDir))
                    {
                        Directory.CreateDirectory(localDir);
                    }

                    // Save the workbook locally in XLSX format
                    workbook.Save(localPath, SaveFormat.Xlsx);
                }

                Console.WriteLine("Workbook loaded from network, compatibility modified, and saved to local path.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
