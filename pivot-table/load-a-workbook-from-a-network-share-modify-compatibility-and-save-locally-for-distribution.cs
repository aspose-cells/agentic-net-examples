// Title: Load Excel workbook from UNC share, enable compatibility & ISO‑29500 strict, and save locally with Aspose.Cells for .NET
// Description: C# sample that loads an Excel file from a network UNC path using Aspose.Cells, activates Settings.CheckCompatibility, applies ISO‑29500:2008 strict compliance, creates the destination folder if needed, and saves the workbook as XLSX for distribution.
// Keywords: Aspose.Cells UNC path | load workbook from network share | Settings.CheckCompatibility | OoxmlCompliance Iso29500_2008_Strict | save workbook locally | C# Aspose.Cells example | create output directory programmatically | fallback workbook if file missing
// Common Searches: Aspose.Cells open Excel from network share C# | Enable compatibility check when saving workbook Aspose.Cells | Set ISO 29500 strict compliance with Aspose.Cells | Save workbook to local folder after loading from UNC | Create missing directory before saving Aspose.Cells
// Developer Intent: Load a workbook from a UNC network location, apply compatibility and strict OOXML settings, and write the file to a local distribution folder.
// Use Cases: Read an existing workbook located at \\ServerName\SharedFolder\SourceWorkbook.xlsx and, if absent, generate a new workbook. | Turn on Settings.CheckCompatibility to ensure backward compatibility with older Excel versions. | Apply Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict for ISO‑29500 strict output. | Automatically create the target directory (e.g., C:\Distribution) before saving the file.
// AI Prompts: Write C# code that opens an Excel file from a UNC path with Aspose.Cells, enables compatibility checking, sets ISO‑29500 strict compliance, creates the output folder if it doesn't exist, and saves the workbook as XLSX. | Explain the impact of Settings.CheckCompatibility and Settings.Compliance on the generated Excel file and how to use them for maximum compatibility. | Provide best‑practice error handling for loading workbooks from network shares using Aspose.Cells in .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsNetworkExample
{
    // C# sample that loads an Excel file from a network UNC path using Aspose.Cells, activates Settings.CheckCompatibility, applies ISO‑29500:2008 strict compliance, creates the destination folder if needed, and saves the workbook as XLSX for distribution.
    class Program
    {
        static void Main()
        {
            // Path to the workbook on a network share (UNC path)
            string networkPath = @"\\ServerName\SharedFolder\SourceWorkbook.xlsx";

            // Local path where the modified workbook will be saved for distribution
            string localPath = @"C:\Distribution\ModifiedWorkbook.xlsx";

            try
            {
                Workbook workbook;

                // Verify that the source workbook exists before attempting to load it
                if (File.Exists(networkPath))
                {
                    // Load the workbook from the network location
                    workbook = new Workbook(networkPath);
                }
                else
                {
                    // If the source file is missing, create a new workbook as a fallback
                    Console.WriteLine($"Source workbook not found at '{networkPath}'. Creating a new workbook.");
                    workbook = new Workbook();
                }

                // Enable compatibility check with earlier Excel versions
                workbook.Settings.CheckCompatibility = true;

                // Set OOXML compliance level (strict ISO compliance)
                workbook.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(localPath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook locally in XLSX format
                workbook.Save(localPath, SaveFormat.Xlsx);

                // Clean up resources
                workbook.Dispose();

                Console.WriteLine("Workbook processed and saved successfully.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
