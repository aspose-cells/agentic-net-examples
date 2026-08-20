// Title: Use a C# using Block to Load and Save an Aspose.Cells Workbook with Automatic Disposal
// Description: This example shows how to create a fallback input.xlsx, open it with Aspose.Cells inside a using statement, modify cell A1, save to output.xlsx, and let the using block call Workbook.Dispose for deterministic resource cleanup.
// Keywords: Aspose.Cells using statement | C# workbook disposal | deterministic resource cleanup | load and save Excel with Aspose | Workbook.Dispose C#
// Common Searches: Aspose.Cells workbook dispose automatically | C# using block for Excel file processing | how to ensure Aspose.Cells releases resources | load, edit, and save Excel with Aspose in C# | create placeholder Excel file if missing Aspose
// Developer Intent: Wrap Aspose.Cells Workbook load and save operations in a using block so the object is disposed automatically.
// Use Cases: Guarantee resource release when processing a single Excel file. | Iterate over many workbooks in a batch job without memory leaks. | Handle missing input files in a web API while ensuring deterministic cleanup.
// AI Prompts: Write C# code that opens an existing Excel file with Aspose.Cells, updates several cells, and saves it, using a using block for disposal. | Refactor a script that creates, loads, and saves workbooks to include a using statement and graceful handling of absent input files. | Generate a method that processes a list of Excel paths, opens each workbook with Aspose.Cells inside a using block, logs errors, and ensures disposal.

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookDemo
{
    // This example shows how to create a fallback input.xlsx, open it with Aspose.Cells inside a using statement, modify cell A1, save to output.xlsx, and let the using block call Workbook.Dispose for deterministic resource cleanup.
    public class WorkbookLoadSaveDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists; create a minimal workbook if it does not.
            if (!File.Exists(inputPath))
            {
                using (Workbook wb = new Workbook())
                {
                    wb.Save(inputPath, SaveFormat.Xlsx);
                }
                Console.WriteLine($"Input file not found. Created empty workbook at '{inputPath}'.");
            }

            // Load the workbook inside a using block for deterministic disposal.
            using (Workbook workbook = new Workbook(inputPath))
            {
                // Example modification: write a value to cell A1 of the first worksheet.
                Worksheet worksheet = workbook.Worksheets[0];
                worksheet.Cells["A1"].PutValue("Updated");

                // Save the workbook while still inside the using block.
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            } // Workbook.Dispose() is called automatically here.
        }
    }
}
