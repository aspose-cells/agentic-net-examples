// Title: C# Load Workbook from FileStream and AutoFill Numeric Series with Aspose.Cells
// Description: This example demonstrates how to open an Excel file (or create a new one if it doesn't exist) using a read‑only FileStream, seed a numeric series in cells A1:A2, auto‑fill the series into a target range with AutoFillType.Series, and save the updated workbook. The code works on .NET platforms and handles missing source files gracefully.
// Keywords: Aspose.Cells | C# | FileStream | load workbook from stream | AutoFill | AutoFillType.Series | CreateRange | numeric series | Excel automation | save workbook
// Common Searches: Aspose.Cells load workbook from filestream c# | C# autofill series using Aspose.Cells | How to use AutoFillType.Series in Aspose.Cells | CreateRange and AutoFill example Aspose.Cells | Save Excel file after AutoFill Aspose.Cells
// Developer Intent: Open an Excel workbook via a FileStream, generate a sequential numeric series by auto‑filling a target range, and persist the changes to a new file.
// Use Cases: Populate a column with sequential IDs in a template that may be missing. | Extend invoice numbers or order counters in a report loaded from a stream. | Generate date or counter series in a newly created workbook before exporting.
// AI Prompts: Show C# code to auto‑fill a numeric series horizontally across columns with Aspose.Cells. | Explain how to set a custom step value when using AutoFillType.Series in Aspose.Cells. | Provide best practices for handling large Excel files loaded from a stream before applying AutoFill.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example demonstrates how to open an Excel file (or create a new one if it doesn't exist) using a read‑only FileStream, seed a numeric series in cells A1:A2, auto‑fill the series into a target range with AutoFillType.Series, and save the updated workbook. The code works on .NET platforms and handles missing source files gracefully.
    public class AutoFillSeriesFromStreamDemo
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
            // Path to the existing Excel file
            string inputPath = "input.xlsx";

            // Path for the resulting Excel file
            string outputPath = "output.xlsx";

            Workbook workbook;

            // Ensure the input file exists; if not, create a new workbook
            if (File.Exists(inputPath))
            {
                // Open the input file as a read‑only stream
                using (FileStream inputStream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
                {
                    // Load the workbook from the stream
                    workbook = new Workbook(inputStream);
                }
            }
            else
            {
                // Create a new workbook with a default worksheet
                workbook = new Workbook();
            }

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Put the first two numbers of the series into the source range (e.g., 1 and 2)
            cells["A1"].PutValue(1);
            cells["A2"].PutValue(2);

            // Define the source range that contains the seed values
            Aspose.Cells.Range sourceRange = cells.CreateRange("A1:A2");

            // Define the target range where the series will be extended (e.g., A3:A10)
            Aspose.Cells.Range targetRange = cells.CreateRange("A3:A10");

            // AutoFill the target range using the Series type to generate a numeric series
            sourceRange.AutoFill(targetRange, AutoFillType.Series);

            // Save the modified workbook to a file
            workbook.Save(outputPath);

            Console.WriteLine("AutoFill operation completed. Result saved to " + Path.GetFullPath(outputPath));
        }
    }
}
