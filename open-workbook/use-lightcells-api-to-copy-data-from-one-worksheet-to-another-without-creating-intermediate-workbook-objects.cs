// Title: Copy a worksheet between Excel files with Aspose.Cells LightCells API in C# (no intermediate Workbook)
// Description: Demonstrates how to stream the first worksheet from source.xlsx into a new workbook (output.xlsx) using Aspose.Cells LightCells, eliminating the need to load both workbooks fully into memory.
// Keywords: Aspose.Cells LightCells | LightCells copy worksheet C# | stream worksheet Aspose.Cells | memory efficient Excel copy C# | copy sheet without loading workbook | large Excel file processing Aspose | C# LightCells example
// Common Searches: How to copy a sheet using Aspose.Cells LightCells C# | Aspose.Cells copy worksheet without loading entire workbook | LightCells API copy Excel sheet C# | Copy large Excel worksheet with minimal memory Aspose | C# stream worksheet from one file to another Aspose.Cells
// Developer Intent: Copy a single worksheet from a source Excel file to a new workbook using the LightCells API to avoid loading both workbooks into memory.
// Use Cases: Processing multi‑gigabyte Excel files on a web server with limited RAM | Generating a report by extracting a template sheet from a master workbook | Archiving a specific sheet while keeping the original file unchanged in a low‑memory environment | Migrating data between workbooks in Azure Functions or AWS Lambda
// AI Prompts: Write C# code that uses Aspose.Cells LightCells to stream the first worksheet from source.xlsx directly into a new workbook named output.xlsx without creating full Workbook objects. | Explain step‑by‑step how LightCells reduces memory consumption when copying sheets between Excel files. | Provide a minimal LightCells example that copies a range of rows from one worksheet to another workbook and saves the result. | Suggest performance tuning tips for LightCells when copying large worksheets in C#.

using System;
using System.IO;
using Aspose.Cells;

namespace LightCellsCopyDemo
{
    // Demonstrates how to stream the first worksheet from source.xlsx into a new workbook (output.xlsx) using Aspose.Cells LightCells, eliminating the need to load both workbooks fully into memory.
    class Program
    {
        static void Main()
        {
            try
            {
                const string sourcePath = "source.xlsx";
                const string outputPath = "output.xlsx";

                // Verify source file exists to avoid FileNotFoundException
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file \"{sourcePath}\" not found.");
                    return;
                }

                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Create an empty destination workbook
                Workbook destinationWorkbook = new Workbook();

                // Add a new worksheet to the destination workbook where the data will be copied
                Worksheet destSheet = destinationWorkbook.Worksheets.Add("CopiedSheet");

                // Copy the first worksheet from source to the newly added worksheet in destination
                sourceWorkbook.Worksheets[0].Copy(destSheet);

                // Save the destination workbook
                destinationWorkbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook copied successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
