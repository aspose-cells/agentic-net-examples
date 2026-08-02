// Title: Transfer Cell Values Only Between Worksheets with Aspose.Cells for .NET
// Description: Demonstrates how to copy just the data from a source range (A1:B2) on one worksheet to a destination range (D5:E6) on another worksheet using Aspose.Cells' CopyValue method. The example populates the source with text, numbers, dates and booleans, applies bold formatting (which is not transferred), and saves the result as an XLSX file.
// Keywords: Aspose.Cells | .NET | C# | CopyValue | copy values only | range copy without formatting | worksheet to worksheet transfer | Excel automation | cell value transfer | Aspose.Cells example
// Common Searches: Aspose.Cells copy only values between worksheets | CopyValue method C# example | Transfer range data without formatting Aspose.Cells | How to copy cell values only in .NET Excel library | Copy values from one sheet to another using Aspose
// Developer Intent: Move the raw values from a source range to a matching destination range on a different worksheet while leaving all formatting untouched.
// Use Cases: Populate a pre‑styled report template with calculation results without altering its design. | Export raw dataset to a new workbook while discarding source cell styles. | Synchronize summary sheets with source data without overwriting existing formatting.
// AI Prompts: Show C# code that copies only the values from a source range to a destination range on another worksheet using Aspose.Cells, preserving destination formatting. | Explain how to copy values from multiple non‑contiguous ranges into a single destination range with Aspose.Cells while keeping existing styles. | Describe handling of mismatched source and destination range sizes when using the CopyValue method in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // Demonstrates how to copy just the data from a source range (A1:B2) on one worksheet to a destination range (D5:E6) on another worksheet using Aspose.Cells' CopyValue method. The example populates the source with text, numbers, dates and booleans, applies bold formatting (which is not transferred), and saves the result as an XLSX file.
    public class TransferValuesOnly
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet as the source sheet and name it
                Worksheet sourceSheet = workbook.Worksheets[0];
                sourceSheet.Name = "Source";

                // Add a second worksheet as the destination sheet and name it
                Worksheet destSheet = workbook.Worksheets[workbook.Worksheets.Add()];
                destSheet.Name = "Destination";

                // Populate some sample data in the source sheet (including different data types)
                sourceSheet.Cells["A1"].PutValue("Text");
                sourceSheet.Cells["A2"].PutValue(123);
                sourceSheet.Cells["A3"].PutValue(45.67);
                sourceSheet.Cells["B1"].PutValue(DateTime.Now);
                sourceSheet.Cells["B2"].PutValue(true);

                // Apply some formatting to source cells (these should NOT be copied)
                Style style = sourceSheet.Cells["A1"].GetStyle();
                style.Font.IsBold = true;
                sourceSheet.Cells["A1"].SetStyle(style);

                // Define the source range (A1:B2) and the destination range (D5:E6)
                AsposeRange srcRange = sourceSheet.Cells.CreateRange("A1:B2");
                AsposeRange destRange = destSheet.Cells.CreateRange("D5:E6");

                // Copy ONLY the values from source range to destination range
                destRange.CopyValue(srcRange);

                // Save the workbook to verify the result
                string outputPath = "TransferValuesOnly.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the console application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
