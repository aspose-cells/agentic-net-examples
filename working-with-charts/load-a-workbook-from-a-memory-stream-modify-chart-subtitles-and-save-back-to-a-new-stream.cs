// Title: Aspose.Cells .NET: Load Workbook from MemoryStream, Update All Chart Subtitles, Save to New Stream
// Description: Demonstrates how to create or load an Excel workbook in memory, iterate through every chart, replace its subtitle with bold 12‑pt text, and write the modified file back to a fresh MemoryStream without touching the file system.
// Keywords: Aspose.Cells | C# chart subtitle | modify chart subtitle | MemoryStream Excel | load workbook from stream | save workbook to stream | chart subtitle formatting | Aspose.Cells chart API | update all charts | Excel chart subtitle .NET
// Common Searches: Aspose.Cells change chart subtitle C# | load Excel from MemoryStream Aspose | update chart subtitle programmatically | save modified workbook to MemoryStream | set chart subtitle font Aspose.Cells | iterate charts in workbook Aspose
// Developer Intent: Replace the subtitle text and styling of every chart in a workbook loaded from a MemoryStream and return the edited workbook as a new stream.
// Use Cases: Generate an Excel report in a web API, adjust chart subtitles on the fly, and stream the result to the client. | Read an uploaded Excel file from a byte array, standardize chart subtitles across all worksheets, and store the updated file back to a database. | Export a modified workbook to a MemoryStream for attaching to an email or uploading to cloud storage without creating temporary files.
// AI Prompts: Write C# code using Aspose.Cells that loads an Excel workbook from a MemoryStream, sets each chart's subtitle to "Updated Subtitle" with bold 12‑pt font, and returns the result as a new MemoryStream. | Show how to correctly dispose of intermediate Workbook and Stream objects after changing chart subtitles with Aspose.Cells. | Explain best‑practice error handling when updating chart subtitles in a workbook read from a stream using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create or load an Excel workbook in memory, iterate through every chart, replace its subtitle with bold 12‑pt text, and write the modified file back to a fresh MemoryStream without touching the file system.
public static class ChartSubtitleModifier
{
    /// <returns>MemoryStream containing the modified workbook in XLSX format.</returns>
    public static MemoryStream ModifyChartSubtitles()
    {
        try
        {
            // -------------------------------------------------
            // 1. Create a sample workbook with data and a chart
            // -------------------------------------------------
            Workbook originalWorkbook = new Workbook(); // uses Workbook() ctor rule
            Worksheet sheet = originalWorkbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";
            chart.Title.Text = "Sample Chart";
            chart.SubTitle.Text = "Original Subtitle";

            // -------------------------------------------------
            // 2. Save the workbook to a memory stream (original)
            // -------------------------------------------------
            MemoryStream originalStream = new MemoryStream();
            originalWorkbook.Save(originalStream, SaveFormat.Xlsx); // uses Save(Stream, SaveFormat) rule
            originalStream.Position = 0; // Reset for reading

            // -------------------------------------------------
            // 3. Load the workbook from the memory stream
            // -------------------------------------------------
            Workbook loadedWorkbook = new Workbook(originalStream); // uses Workbook(Stream) ctor rule

            // -------------------------------------------------
            // 4. Update subtitles of all charts in the workbook
            // -------------------------------------------------
            foreach (Worksheet ws in loadedWorkbook.Worksheets)
            {
                foreach (Chart ch in ws.Charts)
                {
                    // Set a new subtitle text
                    ch.SubTitle.Text = "Updated Subtitle";
                    // Optional: customize subtitle appearance
                    ch.SubTitle.Font.IsBold = true;
                    ch.SubTitle.Font.Size = 12;
                }
            }

            // -------------------------------------------------
            // 5. Save the modified workbook to a new memory stream
            // -------------------------------------------------
            MemoryStream modifiedStream = new MemoryStream();
            loadedWorkbook.Save(modifiedStream, SaveFormat.Xlsx); // uses Save(Stream, SaveFormat) rule
            modifiedStream.Position = 0; // Reset for consumer use

            // Clean up intermediate objects
            originalWorkbook.Dispose();
            loadedWorkbook.Dispose();
            originalStream.Dispose();

            return modifiedStream; // Caller receives the stream containing the updated workbook
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error modifying chart subtitles: {ex.Message}");
            throw;
        }
    }
}

public class Program
{
    public static void Main()
    {
        try
        {
            using (MemoryStream resultStream = ChartSubtitleModifier.ModifyChartSubtitles())
            {
                // Save the resulting workbook to a file for verification
                const string outputPath = "ModifiedChart.xlsx";
                using (FileStream fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                {
                    resultStream.CopyTo(fileStream);
                }

                Console.WriteLine($"Workbook with updated chart subtitles saved to '{outputPath}'.");
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
