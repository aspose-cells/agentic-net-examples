using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaProtectionDemo
{
    public class Program
    {
        public static void Main()
        {
            // Path to the workbook on a network share (UNC path)
            string inputPath = @"\\Server\Share\SampleWorkbook.xlsm";
            // Path where the protected workbook will be saved
            string outputPath = @"\\Server\Share\SampleWorkbook_Protected.xlsm";

            try
            {
                // Verify that the input file exists before attempting to load it
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Ensure the output directory exists
                string? outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Load the workbook from the network share
                Workbook workbook = new Workbook(inputPath);

                // Protect the VBA project if it exists (lock for viewing = false, set password)
                VbaProject? vbaProject = workbook.VbaProject;
                if (vbaProject != null)
                {
                    vbaProject.Protect(false, "MyVbaPassword");
                }
                else
                {
                    Console.WriteLine("No VBA project found in the workbook.");
                }

                // Save the workbook as a macro‑enabled file
                workbook.Save(outputPath, SaveFormat.Xlsm);

                // Reload the saved workbook to verify protection status
                Workbook reloadedWorkbook = new Workbook(outputPath);
                VbaProject? reloadedVba = reloadedWorkbook.VbaProject;

                // Output verification results
                if (reloadedVba != null)
                {
                    Console.WriteLine("VBA Project IsProtected: " + reloadedVba.IsProtected);
                }
                else
                {
                    Console.WriteLine("No VBA project found after saving.");
                }
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}