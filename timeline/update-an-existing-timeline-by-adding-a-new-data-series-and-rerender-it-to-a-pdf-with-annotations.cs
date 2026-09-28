// Title: Add a new data series to an existing timeline chart, insert a worksheet comment, and export the workbook as a PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Load an Excel workbook, append a new NSeries to the first timeline chart using the ranges A2:A10 and B2:B10, add a comment at cell C12, and save the result as a PDF with Aspose.Cells in C#. | Programmatically update a timeline chart by adding a series, place a blue‑styled comment near the chart, and generate a PDF output using Aspose.Cells for .NET.
// Common Searches: C# Aspose.Cells add series to existing chart and save as PDF | How to insert a comment near a chart in an Excel file using Aspose.Cells .NET | Update timeline chart data range programmatically with Aspose.Cells | Export Excel workbook with chart and comments to PDF using Aspose.Cells for .NET | Aspose.Cells add new NSeries to chart from specific cell ranges
// Tags: add NSeries to timeline chart Aspose.Cells | export workbook to PDF with Aspose.Cells | insert worksheet comment Aspose.Cells C# | update chart data range Aspose.Cells | save Excel as PDF including annotations Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

// The C# program loads Timeline.xlsx, adds a new series to the first chart using ranges A2:A10 (categories) and B2:B10 (values), creates a blue comment at cell C12, and saves the updated workbook as Timeline_Updated.pdf in PDF format.
class Program
{
    static void Main()
    {
        try
        {
            const string inputFile = "Timeline.xlsx";
            const string outputFile = "Timeline_Updated.pdf";

            // Verify that the source workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: The file '{inputFile}' was not found.");
                return;
            }

            // Load the existing workbook that contains the timeline chart.
            Workbook workbook = new Workbook(inputFile);

            // Get the first worksheet (adjust the index if necessary).
            Worksheet sheet = workbook.Worksheets[0];

            // Assume the timeline chart is the first chart on the sheet.
            Chart timelineChart = sheet.Charts[0];

            // Define the data ranges for the new series.
            string categoryRange = "A2:A10"; // X‑axis (categories)
            string valueRange = "B2:B10";    // Y‑axis (values)

            // Add the new series to the chart.
            // Combine category and value ranges in a single string.
            int newSeriesIndex = timelineChart.NSeries.Add($"{categoryRange},{valueRange}", true);
            timelineChart.NSeries[newSeriesIndex].Name = "New Series";

            // -----------------------------------------------------------------
            // Add an annotation (comment) to the worksheet near the chart.
            // -----------------------------------------------------------------
            try
            {
                // Add a comment at cell C12 – change the address as needed.
                int commentIndex = sheet.Comments.Add("C12");
                Comment newComment = sheet.Comments[commentIndex];
                newComment.Note = "Added new series to timeline.";
                newComment.Author = "Automation";

                // Optional formatting for the comment text.
                newComment.Font.Color = Color.Blue;
                newComment.Font.Size = 10;
            }
            catch (Exception commentEx)
            {
                Console.WriteLine($"Warning: Unable to add comment. {commentEx.Message}");
            }

            // -----------------------------------------------------------------
            // Save the updated workbook as a PDF, which will include the modified chart and the comment.
            // -----------------------------------------------------------------
            workbook.Save(outputFile, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved successfully as '{outputFile}'.");
        }
        catch (Exception ex)
        {
            // Log or display the exception details for troubleshooting.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
