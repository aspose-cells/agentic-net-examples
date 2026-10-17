// Title: Update Excel built‑in document properties and save the first worksheet as CSV using Aspose.Cells for .NET
// Common Searches: how to change the Title and Author built‑in properties of an Excel file with Aspose.Cells C# | Aspose.Cells C# export first worksheet to CSV after modifying document properties | load existing workbook or create a new one then save as CSV using Aspose.Cells | set Excel metadata programmatically and convert to CSV in .NET | Aspose.Cells save workbook as CSV while preserving built‑in properties
// Tags: Aspose.Cells set built‑in properties | Aspose.Cells conditional workbook loading | Aspose.Cells save worksheet as CSV | C# update Excel metadata title author | Aspose.Cells ensure output directory

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing workbook from "input.xlsx" if it exists, otherwise creates a new workbook, updates the built‑in Title and Author properties, creates the output folder when needed, and saves the first worksheet as a CSV file to "output.csv" using Aspose.Cells for .NET, with error handling.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.csv";

                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new workbook
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook(); // creates a default workbook with one worksheet
                }

                // Update built‑in document properties
                workbook.BuiltInDocumentProperties.Title = "Updated Report Title";
                workbook.BuiltInDocumentProperties.Author = "John Doe";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the first worksheet as CSV
                workbook.Save(outputPath, SaveFormat.Csv);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
