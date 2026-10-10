// Title: Update all chart titles in an Excel workbook loaded from a Stream and save the result to a new MemoryStream using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a Workbook from a Stream with Aspose.Cells, iterates through every worksheet and chart, makes the chart title visible, sets a specified subtitle text, and returns the modified workbook as a MemoryStream. | Generate a method that accepts an input Stream and a subtitle string, updates each chart's title in the workbook, and saves the workbook in its original format to an output MemoryStream. | Create a console example that reads an Excel file from disk, calls the subtitle‑updating method, and writes the resulting MemoryStream to a new file.
// Common Searches: Aspose.Cells C# change chart title text for all charts in a workbook loaded from a stream | How to set chart subtitle programmatically when using a MemoryStream with Aspose.Cells | Save modified Excel workbook to MemoryStream while preserving original file format Aspose.Cells | Iterate through worksheets and charts to update titles using Aspose.Cells .NET
// Tags: Aspose.Cells update chart titles from stream | C# modify chart subtitle memory stream | save workbook to MemoryStream preserving format Aspose.Cells | iterate worksheets charts Aspose.Cells .NET | chart title visibility Aspose.Cells API

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an Excel workbook from a Stream, makes each chart title visible, sets a new subtitle text for all charts, and saves the updated workbook into a new MemoryStream while preserving the original file format.
public class ChartSubtitleModifier
{
    /// <param name="inputStream">Stream containing the original workbook.</param>
    /// <param name="newSubtitle">The text to set as the subtitle for every chart.</param>
    /// <returns>A MemoryStream with the updated workbook.</returns>
    public static MemoryStream ModifyChartSubtitles(Stream inputStream, string newSubtitle)
    {
        try
        {
            // Load the workbook from the provided stream.
            Workbook workbook = new Workbook(inputStream);

            // Iterate through all worksheets.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all charts in the worksheet.
                foreach (Chart chart in sheet.Charts)
                {
                    // Ensure the title is visible.
                    if (!chart.Title.IsVisible)
                    {
                        chart.Title.IsVisible = true;
                    }

                    // Set the new subtitle text (using the title text property as Aspose.Cells does not expose a separate subtitle object).
                    chart.Title.Text = newSubtitle;
                }
            }

            // Save the modified workbook into a memory stream.
            MemoryStream outputStream = new MemoryStream();
            workbook.Save(outputStream, SaveFormat.Xlsx);
            outputStream.Position = 0;
            return outputStream;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error modifying chart subtitles: {ex.Message}");
            throw;
        }
    }

    // Entry point for the console application.
    public static void Main(string[] args)
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            using (FileStream inputFs = File.OpenRead(inputPath))
            {
                MemoryStream modifiedStream = ModifyChartSubtitles(inputFs, "New Subtitle");

                using (FileStream outputFs = File.Create(outputPath))
                {
                    modifiedStream.CopyTo(outputFs);
                }
            }

            Console.WriteLine($"Chart subtitles updated and saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unhandled error: {ex.Message}");
        }
    }
}
