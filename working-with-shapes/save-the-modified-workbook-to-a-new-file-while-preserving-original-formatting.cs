// Title: Save a modified Excel workbook as a new .xlsx file while preserving all original formatting using Aspose.Cells for .NET
// AI Prompts: Load an existing .xlsx file, update a cell value, and employ OoxmlSaveOptions to write the workbook to a separate file while retaining all styles and layout. | Generate C# code that copies an Excel workbook, applies changes, and saves it to a new .xlsx file without losing any original formatting using Aspose.Cells.
// Common Searches: Aspose.Cells .NET how to save a modified workbook to a new file without losing formatting | C# copy Excel file, change a cell, preserve styles using Aspose.Cells | use OoxmlSaveOptions to keep original formatting when saving an Excel workbook in C# | save workbook as new .xlsx while retaining cell styles Aspose.Cells example
// Tags: save workbook with OoxmlSaveOptions .NET | preserve Excel formatting when copying file Aspose.Cells | modify cell and write new .xlsx using Aspose.Cells | Aspose.Cells workbook save preserving styles | C# Aspose.Cells save to new file with original layout

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsExample
{
    // // Loads 'original.xlsx', updates cell A1, and saves the workbook as 'modified.xlsx' using OoxmlSaveOptions to retain all original formatting.
    class Program
    {
        static void Main(string[] args)
        {
            // Define source and destination file paths
            string sourceFile = "original.xlsx";
            string destinationFile = "modified.xlsx";

            try
            {
                // Verify that the source file exists to avoid FileNotFoundException
                if (!File.Exists(sourceFile))
                {
                    Console.WriteLine($"Source file not found: {sourceFile}");
                    return;
                }

                // Load the original workbook
                Workbook workbook = new Workbook(sourceFile);

                // Example modification: change a cell value
                workbook.Worksheets[0].Cells["A1"].PutValue("Updated Value");

                // Prepare save options for .xlsx format
                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx);

                // Save the modified workbook to a new file
                workbook.Save(destinationFile, saveOptions);
                Console.WriteLine($"Workbook saved successfully to {destinationFile}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
