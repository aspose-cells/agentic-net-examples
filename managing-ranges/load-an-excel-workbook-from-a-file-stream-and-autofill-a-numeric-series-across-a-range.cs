// Title: Load Excel workbook from a FileStream and AutoFill a numeric series using Aspose.Cells for .NET (C#)
// Description: This example demonstrates how to open an existing Excel file (or create a blank one) via a FileStream, write initial values to A1 and A2, define source and target ranges, apply AutoFill with AutoFillType.Series to extend the numeric sequence, and save the updated workbook to a new file.
// Keywords: Aspose.Cells | C# | .NET | FileStream Excel load | load workbook from stream | AutoFill range | AutoFillType.Series | numeric series | Excel range creation | save workbook | Excel automation
// Common Searches: Aspose.Cells load workbook from FileStream C# | AutoFill numeric series Aspose.Cells example | How to use AutoFillType.Series in Aspose.Cells | Create and fill Excel range with Aspose.Cells .NET | Read and write Excel file using stream Aspose.Cells
// Developer Intent: Open an Excel workbook from a stream and automatically extend a numeric series across a specified range.
// Use Cases: Programmatically populate sequential IDs in a column after loading a template workbook. | Generate a series of dates or numbers for reporting without manual Excel interaction. | Create a data import routine that reads a file stream, sets a seed range, and fills the rest of the column automatically.
// AI Prompts: Generate C# code that opens an Excel file from a FileStream, writes 1 and 2 to A1:A2, and uses AutoFill to continue the series to A10 with Aspose.Cells. | Explain the behavior of AutoFillType.Series when extending a numeric range in Aspose.Cells. | Show how to modify the sample to AutoFill a date series instead of numbers.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // This example demonstrates how to open an existing Excel file (or create a blank one) via a FileStream, write initial values to A1 and A2, define source and target ranges, apply AutoFill with AutoFillType.Series to extend the numeric sequence, and save the updated workbook to a new file.
    public class LoadAndAutoFillSeries
    {
        // Entry point for the application
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

            // Ensure the input file exists; if not, create a new workbook
            if (!File.Exists(inputPath))
            {
                // Create a blank workbook and save it as the input file
                var newWb = new Workbook();
                newWb.Save(inputPath);
                Console.WriteLine($"Input file not found. Created a new blank workbook at '{inputPath}'.");
            }

            // Open the file as a stream
            using (FileStream stream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
            {
                // Load workbook from the stream
                Workbook workbook = new Workbook(stream);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Put initial numeric values in the source range (e.g., A1:A2)
                cells["A1"].PutValue(1);
                cells["A2"].PutValue(2);

                // Define source and target ranges
                AsposeRange sourceRange = cells.CreateRange("A1:A2");
                AsposeRange targetRange = cells.CreateRange("A3:A10");

                // AutoFill the target range as a series (extends the numeric series)
                sourceRange.AutoFill(targetRange, AutoFillType.Series);

                // Save the modified workbook to a new file
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
        }
    }
}
