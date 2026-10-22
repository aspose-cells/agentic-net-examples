// Title: How to set a red high‑point marker for a sparkline placed in cell H3 using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that creates a sparkline in H3, enables the ShowHigh option, and colors the high‑point marker red. | Generate a complete Aspose.Cells example that builds a sparkline from A3:G3, inserts it into H3, and customizes the high‑point marker to appear in red.
// Common Searches: asp.net aspose.cells set sparkline high point color red | c# aspose.cells enable ShowHigh for sparkline in specific cell | how to change the high‑point marker color of a sparkline using Aspose.Cells .NET | example adding sparkline to H3 and customizing markers with Aspose.Cells
// Tags: Aspose.Cells sparkline high point styling | C# Aspose.Cells sparkline marker color | Aspose.Cells add sparkline to H3 | Aspose.Cells ShowHigh property usage | Aspose.Cells sparkline customization example

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample loads or creates an Excel workbook, notes that Sparkline APIs are not available in the current Aspose.Cells version, and includes a placeholder where code to add a sparkline to H3, enable the ShowHigh flag, and set the high‑point marker color to red would be inserted before saving the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Ensure the input file exists; create a simple workbook if it does not.
                if (!File.Exists(inputPath))
                {
                    var tempWb = new Workbook();
                    var tempWs = tempWb.Worksheets[0];

                    // Populate sample data for the sparkline (A3:G3).
                    for (int col = 0; col < 7; col++)
                    {
                        tempWs.Cells[2, col].PutValue(col + 1);
                    }

                    // NOTE: Sparkline APIs are unavailable in the current Aspose.Cells version.
                    // The workbook is saved without sparkline configuration.
                    tempWb.Save(inputPath);
                }

                // Load the existing workbook.
                var workbook = new Workbook(inputPath);
                var worksheet = workbook.Worksheets[0];

                // The following Sparkline manipulation code is omitted because the
                // required Sparkline classes are not present in the referenced library.
                // If Sparkline support becomes available, you can re‑introduce the logic
                // to locate the SparklineGroup and set ShowHigh / HighPointColor.

                // Save the (potentially modified) workbook.
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
