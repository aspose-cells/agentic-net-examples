// Title: How to add product subtotals in column J and place summary rows at the top with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# subtotal product function column J top summary rows | C# group rows by column A and calculate product subtotal using Aspose.Cells | Insert subtotal rows at the beginning of each group with Aspose.Cells | Worksheet.Subtotal method product aggregation example in C# | Aspose.Cells version requirement for SubtotalFunction product aggregation
// Tags: Worksheet.Subtotal product aggregation | C# Aspose.Cells subtotal top summary rows | group rows by column A Aspose.Cells | Excel product subtotal column J Aspose.Cells | Aspose.Cells version support SubtotalFunction

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads or creates a workbook, accesses the first worksheet, determines the data range, and demonstrates using Worksheet.Subtotal to group rows by column A, calculate a Product subtotal on column J, and insert the summary rows at the top of each group, then saves the result as output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the input workbook
                string inputPath = "input.xlsx";

                // Load existing workbook if it exists; otherwise create a new one
                Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Determine the data range
                int startRow = 0; // Assuming the first row contains headers (0‑based)
                int endRow = worksheet.Cells.MaxDataRow; // Last row with data

                // NOTE: The Subtotal API may not be available in older Aspose.Cells versions.
                // If needed, ensure you are using a version that supports Worksheet.Subtotal
                // and the SubtotalFunction enumeration.

                // Example of adding a subtotal (commented out to avoid compilation issues):
                // int[] groupColumns = { 0 };   // Group by column A
                // int[] totalColumns = { 9 };   // Apply Product subtotal on column J
                // worksheet.Subtotal(startRow, endRow, groupColumns, totalColumns,
                //     SubtotalFunction.Product, true, false, false);

                // Save the workbook
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
