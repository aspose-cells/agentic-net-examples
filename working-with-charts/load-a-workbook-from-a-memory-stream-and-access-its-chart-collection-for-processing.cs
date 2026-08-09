// Title: Load Workbook from MemoryStream and Manipulate Chart Collection with Aspose.Cells for .NET
// Description: Shows how to create a workbook containing a column chart, save it to a MemoryStream, reload it using the Workbook(Stream) constructor, iterate the first worksheet's ChartCollection, prepend text to each chart title, and save the modified file.
// Keywords: Aspose.Cells | MemoryStream | load workbook from stream | chart collection | modify chart title | .NET | C# | Excel chart processing | in‑memory workbook | SaveToStream | Workbook(Stream) ctor
// Common Searches: Aspose.Cells load workbook from memory stream | access charts after loading workbook from stream | change chart title C# Aspose.Cells | enumerate chart collection Aspose.Cells | reset MemoryStream.Position before Workbook ctor | process Excel charts in memory
// Developer Intent: Read an Excel file from a stream, edit its charts, and write the updated workbook.
// Use Cases: Receive an Excel file as a byte array, load it via MemoryStream, add a branding prefix to every chart title, and return the revised file. | Load a template workbook from a stream, switch chart types or data series programmatically, then export the final report. | Batch‑process multiple in‑memory workbooks, applying uniform chart formatting before persisting each file.
// AI Prompts: Write C# code that uses Aspose.Cells to open a workbook from a MemoryStream, loop through all charts, and prepend a custom string to each chart title. | Explain why resetting MemoryStream.Position to 0 is required before constructing a Workbook with Aspose.Cells. | Show how to clone charts from one workbook into another after loading the source from a stream, then save the combined workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartProcessing
{
    // Shows how to create a workbook containing a column chart, save it to a MemoryStream, reload it using the Workbook(Stream) constructor, iterate the first worksheet's ChartCollection, prepend text to each chart title, and save the modified file.
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Create a sample workbook with a chart (source workbook)
            // ------------------------------------------------------------
            Workbook sourceWorkbook = new Workbook();

            // Populate sample data
            Worksheet srcSheet = sourceWorkbook.Worksheets[0];
            srcSheet.Cells["A1"].PutValue("Category");
            srcSheet.Cells["A2"].PutValue("A");
            srcSheet.Cells["A3"].PutValue("B");
            srcSheet.Cells["A4"].PutValue("C");
            srcSheet.Cells["B1"].PutValue("Value");
            srcSheet.Cells["B2"].PutValue(10);
            srcSheet.Cells["B3"].PutValue(20);
            srcSheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIdx = srcSheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart srcChart = srcSheet.Charts[chartIdx];
            srcChart.NSeries.Add("B2:B4", true);
            srcChart.NSeries.CategoryData = "A2:A4";
            srcChart.Title.Text = "Sample Chart";

            // ------------------------------------------------------------
            // 2. Save the source workbook to a memory stream (using SaveToStream)
            // ------------------------------------------------------------
            MemoryStream memoryStream = sourceWorkbook.SaveToStream();

            // Reset the stream position to the beginning before reading
            memoryStream.Position = 0;

            // ------------------------------------------------------------
            // 3. Load a new workbook from the memory stream (using Workbook(Stream) ctor)
            // ------------------------------------------------------------
            Workbook loadedWorkbook = new Workbook(memoryStream);

            // ------------------------------------------------------------
            // 4. Access the chart collection of the first worksheet
            // ------------------------------------------------------------
            Worksheet firstSheet = loadedWorkbook.Worksheets[0];
            ChartCollection charts = firstSheet.Charts;

            // ------------------------------------------------------------
            // 5. Process each chart (example: output info and modify title)
            // ------------------------------------------------------------
            for (int i = 0; i < charts.Count; i++)
            {
                Chart chart = charts[i];
                Console.WriteLine($"Chart {i}: Type = {chart.Type}, Name = {chart.Name}");

                // Example modification: prepend "Processed - " to the chart title
                if (!string.IsNullOrEmpty(chart.Title.Text))
                {
                    chart.Title.Text = "Processed - " + chart.Title.Text;
                }
            }

            // ------------------------------------------------------------
            // 6. Save the processed workbook to a file
            // ------------------------------------------------------------
            loadedWorkbook.Save("ProcessedWorkbook.xlsx");

            // Clean up
            memoryStream.Dispose();
            sourceWorkbook.Dispose();
            loadedWorkbook.Dispose();

            Console.WriteLine("Workbook loaded from memory stream, charts processed, and saved as 'ProcessedWorkbook.xlsx'.");
        }
    }
}
