// Title: Modify Excel Chart Subtitle from a MemoryStream and Save to a New Stream with Aspose.Cells (C#)
// Description: Creates a workbook with a column chart, saves it to a MemoryStream (XLS), reloads it, updates the chart subtitle text and style, and writes the result to a new MemoryStream (XLSX) using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | MemoryStream | load workbook from stream | save workbook to stream | chart subtitle | chart subtitle formatting | update chart title | Excel chart manipulation | in‑memory Excel processing
// Common Searches: Aspose.Cells change chart subtitle from MemoryStream | C# load Excel file from byte array and edit chart | save modified Excel workbook to new MemoryStream Aspose | update chart subtitle font size Aspose.Cells C# | convert XLS to XLSX after editing chart subtitle
// Developer Intent: Load an Excel workbook from a MemoryStream, modify the chart subtitle text and formatting, and save the updated workbook to another MemoryStream.
// Use Cases: Edit chart subtitles in Excel files received as byte arrays before sending them to a client. | Apply uniform subtitle styling to all charts in dynamically generated reports stored in memory. | Convert an in‑memory XLS workbook to XLSX after programmatically updating chart titles and subtitles.
// AI Prompts: Write C# code that loads an Excel workbook from a MemoryStream, sets every chart subtitle to bold 14pt text, and returns the workbook as a new MemoryStream using Aspose.Cells. | Show how to iterate through all charts in a workbook loaded from a stream and set each subtitle to "Report Generated" with italic style.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook with a column chart, saves it to a MemoryStream (XLS), reloads it, updates the chart subtitle text and style, and writes the result to a new MemoryStream (XLSX) using Aspose.Cells for .NET.
class ChartSubtitleExample
{
    static void Main()
    {
        // 1. Create a workbook with sample data and a chart
        Workbook wb = new Workbook();                                   // Workbook()
        Worksheet ws = wb.Worksheets[0];
        ws.Cells["A1"].PutValue("Category");
        ws.Cells["A2"].PutValue("A");
        ws.Cells["A3"].PutValue("B");
        ws.Cells["A4"].PutValue("C");
        ws.Cells["B1"].PutValue("Value");
        ws.Cells["B2"].PutValue(10);
        ws.Cells["B3"].PutValue(20);
        ws.Cells["B4"].PutValue(30);

        int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 20, 8);    // add chart
        Chart chart = ws.Charts[chartIdx];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";
        chart.Title.Text = "Main Title";
        chart.SubTitle.Text = "Original Subtitle";                      // set initial subtitle

        // 2. Save the workbook to a memory stream (xls format)
        MemoryStream sourceStream = wb.SaveToStream();                  // Workbook.SaveToStream()
        sourceStream.Position = 0;                                      // reset for reading

        // 3. Load the workbook from the memory stream
        Workbook loadedWb = new Workbook(sourceStream);                  // Workbook(Stream)
        Worksheet loadedWs = loadedWb.Worksheets[0];
        Chart loadedChart = loadedWs.Charts[0];

        // 4. Modify the chart subtitle
        loadedChart.SubTitle.Text = "Updated Subtitle";
        loadedChart.SubTitle.Font.IsBold = true;
        loadedChart.SubTitle.Font.Size = 12;

        // 5. Save the modified workbook to a new memory stream (xlsx format)
        MemoryStream resultStream = new MemoryStream();
        loadedWb.Save(resultStream, SaveFormat.Xlsx);                   // Workbook.Save(Stream, SaveFormat)
        resultStream.Position = 0;                                      // ready for further use

        // Optional: verify that the stream contains data
        Console.WriteLine($"Result stream length: {resultStream.Length} bytes");
    }
}
