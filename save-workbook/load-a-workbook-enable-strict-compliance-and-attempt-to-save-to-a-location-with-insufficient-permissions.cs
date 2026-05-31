using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class StrictComplianceSaveDemo
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            // Path to an existing workbook to load
            string inputPath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            Workbook workbook = null;
            try
            {
                // Load the workbook
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Enable ISO/IEC 29500:2008 Strict compliance for OOXML
            workbook.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

            // Destination path with likely insufficient write permissions
            string protectedPath = @"C:\Windows\System32\protected_output.xlsx";

            try
            {
                // Attempt to save the workbook to the protected location
                workbook.Save(protectedPath);
                Console.WriteLine("Workbook saved successfully (unexpected).");
            }
            catch (Exception ex)
            {
                // Expected failure due to insufficient permissions
                Console.WriteLine("Failed to save workbook: " + ex.Message);
            }
        }
    }
}