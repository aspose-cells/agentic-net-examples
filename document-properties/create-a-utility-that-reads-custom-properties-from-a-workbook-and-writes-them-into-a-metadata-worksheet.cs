// Title: Read custom document properties from an Excel workbook and export them to a "Metadata" worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an existing .xlsx file with Aspose.Cells, iterates over its CustomDocumentProperties collection, creates (or clears) a worksheet named "Metadata", and writes each property's name and value into two columns. | Extend the program to also list built‑in document properties (Author, Title, etc.) together with custom properties in the same metadata sheet. | Add functionality that saves the generated "Metadata" worksheet as a separate CSV file while leaving the original workbook unchanged.
// Common Searches: aspnet read custom document properties from Excel using Aspose.Cells | c# generate a metadata worksheet with workbook properties Aspose.Cells | how to export Excel custom properties to a new worksheet in .NET | save custom document properties to CSV with Aspose.Cells C#
// Tags: metadata extraction from Excel using Aspose.Cells | generate metadata worksheet in .xlsx with C# | populate worksheet cells with property values Aspose.Cells | auto‑fit columns after writing data Aspose.Cells | export worksheet to CSV with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookMetadataUtility
{
    // The utility loads an existing Excel file, retrieves its custom document properties, creates or clears a worksheet named "Metadata", writes each property's name and value into two columns, auto‑fits the columns for readability, and saves the updated workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for source and output workbooks
            string sourcePath = "InputWorkbook.xlsx";
            string outputPath = "OutputWorkbook.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(sourcePath);

                // Access custom document properties collection
                var customProps = workbook.CustomDocumentProperties;

                // Find existing "Metadata" worksheet or create a new one
                Worksheet metadataSheet = null;
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    if (ws.Name.Equals("Metadata", StringComparison.OrdinalIgnoreCase))
                    {
                        metadataSheet = ws;
                        break;
                    }
                }

                if (metadataSheet == null)
                {
                    // Add a new worksheet named "Metadata"
                    int sheetIndex = workbook.Worksheets.Add();
                    metadataSheet = workbook.Worksheets[sheetIndex];
                    metadataSheet.Name = "Metadata";
                }
                else
                {
                    // Clear all existing cells in the sheet
                    metadataSheet.Cells.Clear();
                }

                // Write header titles
                metadataSheet.Cells[0, 0].PutValue("Property Name");
                metadataSheet.Cells[0, 1].PutValue("Value");

                // Populate the sheet with custom property name/value pairs
                int rowIndex = 1; // Start after header row
                foreach (var prop in customProps)
                {
                    // Property name
                    metadataSheet.Cells[rowIndex, 0].PutValue(prop.Name);
                    // Property value (convert to string for safety)
                    metadataSheet.Cells[rowIndex, 1].PutValue(prop.Value?.ToString() ?? string.Empty);
                    rowIndex++;
                }

                // Auto‑fit columns for better readability
                metadataSheet.AutoFitColumns();

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook with the new "Metadata" sheet
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime exceptions gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
