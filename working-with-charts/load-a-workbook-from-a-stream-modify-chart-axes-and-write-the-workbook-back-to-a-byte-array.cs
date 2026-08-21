// Title: Modify Excel Chart Axes from a Byte Array with Aspose.Cells for .NET
// Description: Loads an Excel workbook from a byte[] using a MemoryStream, accesses the first worksheet and its first chart, updates the category and value axis titles, sets major tick marks to outside, optionally defines the value axis range, then saves the workbook back to a byte[] in XLSX format.
// Keywords: Aspose.Cells | C# | .NET | chart axis modification | load workbook from byte array | save workbook to byte array | Excel chart formatting | memory stream | set axis titles | tick mark style | axis range
// Common Searches: Aspose.Cells change chart axis title programmatically | load Excel file from byte[] and edit chart | save modified Excel workbook to byte array C# | set chart tick marks Aspose.Cells .NET | adjust chart axis range using Aspose.Cells
// Developer Intent: Edit the axis properties of the first chart in a workbook that is read from a byte array and return the updated workbook as a byte array.
// Use Cases: Server‑side processing of uploaded Excel files to rename chart axes before sending the file back to the client. | Dynamic report generation where axis titles and ranges are calculated from data and the result is streamed via a web API. | Document conversion pipelines that need to modify chart formatting without writing intermediate files to disk.
// AI Prompts: Create a C# method that receives a byte[] of an Excel file, changes the first chart's category and value axis titles, sets major tick marks to outside, defines a value axis range of 0‑100, and returns the modified workbook as a byte[]. | Add robust error handling to the chart‑axis modification routine: check for the existence of charts, log a warning if none are found, and handle unexpected exceptions gracefully. | Extend the ModifyChartAxes function to iterate through all worksheets and update every chart's axis titles and ranges based on supplied parameters.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // Loads an Excel workbook from a byte[] using a MemoryStream, accesses the first worksheet and its first chart, updates the category and value axis titles, sets major tick marks to outside, optionally defines the value axis range, then saves the workbook back to a byte[] in XLSX format.
    public class ChartAxisModifier
    {
        /// <param name="inputExcel">The input Excel file bytes.</param>
        /// <returns>Byte array of the modified workbook.</returns>
        public static byte[] ModifyChartAxes(byte[] inputExcel)
        {
            // Load the workbook from the input stream
            using (MemoryStream inputStream = new MemoryStream(inputExcel))
            {
                Workbook workbook = new Workbook(inputStream);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure there is at least one chart to modify
                if (worksheet.Charts.Count > 0)
                {
                    // Get the first chart
                    Chart chart = worksheet.Charts[0];

                    // Modify axis titles
                    chart.CategoryAxis.Title.Text = "New Category Axis";
                    chart.ValueAxis.Title.Text = "New Value Axis";

                    // Set major tick marks for both axes
                    chart.CategoryAxis.MajorTickMark = TickMarkType.Outside;
                    chart.ValueAxis.MajorTickMark = TickMarkType.Outside;

                    // Optionally set axis range (example values)
                    chart.ValueAxis.MinValue = 0;
                    chart.ValueAxis.MaxValue = 100;
                }

                // Save the modified workbook to a memory stream as XLSX
                using (MemoryStream outputStream = new MemoryStream())
                {
                    workbook.Save(outputStream, SaveFormat.Xlsx);
                    return outputStream.ToArray();
                }
            }
        }
    }

    public class Program
    {
        // Entry point required for compilation
        public static void Main(string[] args)
        {
            try
            {
                // Determine input and output file paths
                string inputPath = args.Length > 0 ? args[0] : "input.xlsx";
                string outputPath = args.Length > 1 ? args[1] : "output_modified.xlsx";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Read input workbook bytes
                byte[] inputBytes = File.ReadAllBytes(inputPath);

                // Modify chart axes
                byte[] resultBytes = ChartAxisModifier.ModifyChartAxes(inputBytes);

                // Write the modified workbook to the output file
                File.WriteAllBytes(outputPath, resultBytes);
                Console.WriteLine($"Modified workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
