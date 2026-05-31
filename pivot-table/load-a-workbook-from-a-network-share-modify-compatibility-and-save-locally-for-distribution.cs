using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsNetworkExample
{
    class Program
    {
        static void Main()
        {
            // Path to the workbook on a network share (UNC path)
            string networkPath = @"\\Server\Share\Documents\source.xlsx";

            // Local path where the modified workbook will be saved for distribution
            string localPath = @"C:\Distribution\output.xlsx";

            Workbook workbook = null;

            try
            {
                // Verify source file exists before loading
                if (!File.Exists(networkPath))
                {
                    Console.WriteLine($"Source file not found: {networkPath}");
                    return;
                }

                // Load the workbook from the network location
                workbook = new Workbook(networkPath);

                // Enable compatibility checks for older Excel versions
                workbook.Settings.CheckCompatibility = true;

                // Set OOXML compliance to strict (ISO/IEC 29500:2008)
                workbook.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(localPath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook locally in XLSX format
                workbook.Save(localPath, SaveFormat.Xlsx);

                Console.WriteLine("Workbook loaded from network, compatibility modified, and saved to local path.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
            finally
            {
                workbook?.Dispose();
            }
        }
    }
}