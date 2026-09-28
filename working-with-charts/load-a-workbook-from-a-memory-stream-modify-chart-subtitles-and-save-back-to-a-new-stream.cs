// Title: Modify all chart subtitles in an Excel workbook loaded from a MemoryStream and save the updated file to a new MemoryStream using Aspose.Cells for .NET
// AI Prompts: Write C# code that reads an XLSX workbook from a MemoryStream, iterates over every worksheet and chart, sets each chart's Title.Text to a given subtitle string, and returns the modified workbook as a new MemoryStream with Aspose.Cells. | Show how to use Aspose.Cells to load an Excel file from a MemoryStream, update chart subtitles, and save the workbook back to another MemoryStream in XLSX format.
// Common Searches: Aspose.Cells C# change chart subtitle in workbook loaded from MemoryStream | How to update all chart titles in an Excel file using Aspose.Cells without saving to disk | Save modified Excel workbook to a new MemoryStream after editing chart properties | Iterate through worksheets and charts in Aspose.Cells to set a common subtitle | Load XLSX from stream, modify chart titles, and get output stream in .NET
// Tags: Aspose.Cells update chart subtitle from MemoryStream | C# iterate worksheets charts Aspose.Cells | save modified workbook to MemoryStream XLSX | chart title text replacement Aspose.Cells | in‑memory Excel processing Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ChartSubtitleModifierApp
{
    // Loads an XLSX workbook from an input MemoryStream, loops through each worksheet and its charts to set the chart Title.Text to the supplied subtitle, then saves the workbook into a new MemoryStream in XLSX format.
    public class ChartSubtitleModifier
    {
        public static MemoryStream ModifyChartSubtitles(MemoryStream inputStream, string newSubtitle)
        {
            try
            {
                // Load workbook from memory stream
                Workbook workbook = new Workbook(inputStream);

                // Iterate through worksheets and charts
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (Chart chart in sheet.Charts)
                    {
                        // Update the chart title (used here as subtitle)
                        if (chart.Title != null)
                        {
                            chart.Title.Text = newSubtitle;
                        }
                    }
                }

                // Save modified workbook to a new memory stream
                MemoryStream outputStream = new MemoryStream();
                workbook.Save(outputStream, SaveFormat.Xlsx);
                outputStream.Position = 0;
                return outputStream;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error modifying chart subtitles: {ex.Message}");
                throw;
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";
            string newSubtitle = "Updated Subtitle";

            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.Error.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load input file into a memory stream
                using (FileStream fs = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
                using (MemoryStream inputMs = new MemoryStream())
                {
                    fs.CopyTo(inputMs);
                    inputMs.Position = 0;

                    // Modify chart subtitles
                    MemoryStream resultMs = ChartSubtitleModifier.ModifyChartSubtitles(inputMs, newSubtitle);

                    // Write the result to the output file
                    using (FileStream outFs = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                    {
                        resultMs.CopyTo(outFs);
                    }

                    Console.WriteLine($"Modified workbook saved to {outputPath}");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
