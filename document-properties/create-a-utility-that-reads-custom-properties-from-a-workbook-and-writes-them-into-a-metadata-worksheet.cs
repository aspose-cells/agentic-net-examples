// Title: Export custom document properties from an Excel workbook to a "Metadata" worksheet using Aspose.Cells for .NET (C#)
// Common Searches: how to read custom document properties from an Excel file and write them to a new sheet using Aspose.Cells C# | Aspose.Cells export custom properties to a worksheet named Metadata in .NET | C# code to create a Metadata worksheet and list workbook custom properties with Aspose.Cells | save workbook with added Metadata sheet while preserving custom properties using Aspose.Cells | ensure output directory exists before saving Excel workbook with Aspose.Cells C#
// Tags: retrieve custom document properties Aspose.Cells | add Metadata sheet to workbook Aspose.Cells | populate Excel cells with property name and value Aspose.Cells | save workbook to new location ensuring directory exists C# | access CustomDocumentProperties collection Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookMetadataApp
{
    // The utility loads an Excel workbook, reads its CustomDocumentProperties collection, creates a worksheet named "Metadata" if missing, writes a header row and then each property's name and value into columns A and B, ensures the destination folder exists, and saves the updated workbook to the specified path, handling any errors with a try‑catch block.
    public class WorkbookMetadataUtility
    {
        /// <param name="sourcePath">Full path to the input workbook.</param>
        /// <param name="destinationPath">Full path where the modified workbook will be saved.</param>
        public static void ExportCustomProperties(string sourcePath, string destinationPath)
        {
            try
            {
                // Verify source file exists to avoid FileNotFoundException.
                if (!File.Exists(sourcePath))
                    throw new FileNotFoundException($"Source file not found: {sourcePath}");

                // Load the workbook from the specified file.
                var workbook = new Workbook(sourcePath);

                // Access the collection of custom document properties.
                var customProps = workbook.CustomDocumentProperties;

                // Retrieve or create the "Metadata" worksheet.
                var metadataSheet = workbook.Worksheets["Metadata"];
                if (metadataSheet == null)
                {
                    int sheetIndex = workbook.Worksheets.Add();
                    metadataSheet = workbook.Worksheets[sheetIndex];
                    metadataSheet.Name = "Metadata";
                }

                // Write header titles.
                metadataSheet.Cells[0, 0].PutValue("Property Name");
                metadataSheet.Cells[0, 1].PutValue("Property Value");

                // Populate the sheet with each custom property's name and value.
                int currentRow = 1;
                foreach (var prop in customProps)
                {
                    metadataSheet.Cells[currentRow, 0].PutValue(prop.Name);
                    metadataSheet.Cells[currentRow, 1].PutValue(prop.Value);
                    currentRow++;
                }

                // Ensure the destination directory exists.
                var destDir = Path.GetDirectoryName(destinationPath);
                if (!string.IsNullOrEmpty(destDir) && !Directory.Exists(destDir))
                {
                    Directory.CreateDirectory(destDir);
                }

                // Save the workbook with the new "Metadata" sheet.
                workbook.Save(destinationPath);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error exporting custom properties: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example usage: adjust paths as needed.
            string sourcePath = "input.xlsx";
            string destinationPath = "output_with_metadata.xlsx";

            WorkbookMetadataUtility.ExportCustomProperties(sourcePath, destinationPath);
        }
    }
}
