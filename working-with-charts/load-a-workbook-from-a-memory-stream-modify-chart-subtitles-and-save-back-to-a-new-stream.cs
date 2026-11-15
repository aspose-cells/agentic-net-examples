// Title: Update all chart titles in an Excel workbook loaded from a MemoryStream and save to a new MemoryStream using Aspose.Cells for .NET
// Common Searches: how to programmatically change all chart titles in an Excel file using Aspose.Cells for .NET | update chart subtitle text when workbook is loaded from a MemoryStream with Aspose.Cells | save modified Excel workbook to a new MemoryStream after editing chart titles in C# | Aspose.Cells example for iterating charts and setting title visibility | C# code to load workbook from stream, modify chart titles, and export to stream
// Tags: load workbook from stream Aspose.Cells .NET | iterate worksheets and charts Aspose.Cells | set chart title visibility Aspose.Cells | assign chart title text programmatically Aspose.Cells | save workbook to memory stream Aspose.Cells | update all chart subtitles Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads an Excel workbook from an input stream, iterates through each worksheet and its charts, makes each chart title visible, sets a new subtitle text, and then saves the updated workbook into a new MemoryStream in XLSX format.
public class ChartSubtitleUpdater
{
    /// <param name="inputStream">Stream containing the original workbook.</param>
    /// <param name="newSubtitle">The text to set as the title for every chart.</param>
    /// <returns>A memory stream with the updated workbook.</returns>
    public static MemoryStream UpdateChartSubtitles(Stream inputStream, string newSubtitle)
    {
        try
        {
            // Load the workbook from the provided stream.
            Workbook workbook = new Workbook(inputStream);

            // Iterate through all worksheets.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all charts on the worksheet.
                foreach (Chart chart in sheet.Charts)
                {
                    // Ensure the title is visible and set its text.
                    chart.Title.IsVisible = true;
                    chart.Title.Text = newSubtitle;
                }
            }

            // Save the modified workbook into a memory stream.
            MemoryStream outputStream = new MemoryStream();
            workbook.Save(outputStream, SaveFormat.Xlsx);
            outputStream.Position = 0; // Reset for reading.

            return outputStream;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error updating chart subtitles: {ex.Message}");
            throw;
        }
    }

    // Entry point for demonstration / testing.
    public static void Main(string[] args)
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";
        string subtitleText = "Updated Subtitle";

        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            using (FileStream inputStream = File.OpenRead(inputPath))
            using (MemoryStream updatedStream = UpdateChartSubtitles(inputStream, subtitleText))
            {
                // Write the updated workbook to disk.
                File.WriteAllBytes(outputPath, updatedStream.ToArray());
                Console.WriteLine($"Workbook saved with updated subtitles to: {outputPath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Processing failed: {ex.Message}");
        }
    }
}
