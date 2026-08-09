// Title: C# – Check and Add Custom Document Property "ClientName" with Aspose.Cells
// Description: Loads an Excel workbook (or creates a new one if the file is missing), examines the CustomDocumentProperties collection, adds a "ClientName" property only when it does not already exist, creates the output directory if needed, and saves the workbook.
// Keywords: Aspose.Cells | C# | CustomDocumentProperties | check custom property | add custom property | Excel metadata | load workbook | save workbook | client name property | Excel automation
// Common Searches: Aspose.Cells check if custom document property exists | C# add custom property to Excel workbook | How to create or open workbook and manage custom properties with Aspose.Cells | Add missing custom property to Excel file using .NET | Aspose.Cells example for conditional custom property insertion
// Developer Intent: Identify whether the workbook already contains a custom property named "ClientName" and insert it only if it is absent.
// Use Cases: Embed client identifiers in report workbooks to maintain traceability without duplicating entries. | Refresh existing Excel templates programmatically while preserving previously set custom properties. | Generate a fresh workbook with a default client name when the source file is unavailable, then continue processing.
// AI Prompts: Generate C# code with Aspose.Cells that opens an Excel file, checks for a custom property called "ProjectId", and adds it with a specified value only if it does not exist. | Refactor the example to accept the property name and value as method parameters and handle a list of properties in a loop. | Explain best practices for exception handling when working with CustomDocumentProperties in Aspose.Cells and ensure the workbook is saved even if the input file is missing.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an Excel workbook (or creates a new one if the file is missing), examines the CustomDocumentProperties collection, adds a "ClientName" property only when it does not already exist, creates the output directory if needed, and saves the workbook.
    public class CheckAndAddCustomProperty
    {
        public static void Run()
        {
            string inputPath = "InputWorkbook.xlsx";
            string outputPath = "OutputWorkbook.xlsx";

            try
            {
                Workbook workbook;

                // Load existing workbook or create a new one if the file is missing
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                    Console.WriteLine($"Input file '{inputPath}' not found. Created a new workbook.");
                }

                // Access custom document properties
                var customProps = workbook.CustomDocumentProperties;

                // Add "ClientName" property if it does not exist
                if (!customProps.Contains("ClientName"))
                {
                    customProps.Add("ClientName", "Acme Corp");
                    Console.WriteLine("Custom property 'ClientName' added.");
                }
                else
                {
                    Console.WriteLine("Custom property 'ClientName' already exists.");
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
