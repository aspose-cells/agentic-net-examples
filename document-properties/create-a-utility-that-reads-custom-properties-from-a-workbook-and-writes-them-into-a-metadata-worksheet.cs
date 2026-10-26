// Title: Create a Metadata worksheet that lists custom document properties from an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells read custom document properties and write them to a new sheet | How to generate a Metadata tab with custom property name/value pairs in an Excel file using Aspose.Cells | Aspose.Cells add a worksheet named Metadata if it does not exist and populate it with custom properties | Export custom document properties to an Excel sheet programmatically with Aspose.Cells .NET | Save workbook with custom properties displayed in a Metadata worksheet using Aspose.Cells
// Tags: Aspose.Cells read custom document properties | Aspose.Cells write properties to worksheet | create Metadata worksheet if missing Aspose.Cells | auto-fit columns after populating data Aspose.Cells | save updated workbook with new sheet .NET

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookMetadataUtility
{
    // The example loads an existing Excel workbook, extracts its custom document properties, ensures a worksheet named "Metadata" exists (creating it when absent), writes each property name and value into two columns with headers, auto‑fits the columns for readability, and saves the modified workbook to a new file while handling missing input files and runtime exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for source and output workbooks
            string inputPath = "SourceWorkbook.xlsx";
            string outputPath = "SourceWorkbook_WithMetadata.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Retrieve custom document properties collection
                var customProps = workbook.CustomDocumentProperties;

                // Ensure a worksheet named "Metadata" exists; create if it does not
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
                    // Add a new worksheet at the end and rename it to "Metadata"
                    int sheetIndex = workbook.Worksheets.Add();
                    metadataSheet = workbook.Worksheets[sheetIndex];
                    metadataSheet.Name = "Metadata";
                }

                // Write headers
                Cells cells = metadataSheet.Cells;
                cells["A1"].PutValue("Property Name");
                cells["B1"].PutValue("Property Value");

                // Populate the sheet with custom property name/value pairs
                int row = 1; // zero‑based index; row 1 is the second row (after headers)
                foreach (var prop in customProps)
                {
                    cells[row, 0].PutValue(prop.Name);
                    cells[row, 1].PutValue(prop.Value);
                    row++;
                }

                // Auto‑fit columns for better readability
                metadataSheet.AutoFitColumns();

                // Save the workbook with the new "Metadata" sheet
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
