// Title: Create a clickable label hyperlink on an Excel chart using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that inserts a label shape onto a worksheet, styles it as a blue underlined link, and attaches a hyperlink that opens a specified URL when the chart is displayed. | Generate a .NET example that creates a workbook, adds a column chart, places a label at a chosen position, sets its text and screen tip, and configures the label’s Hyperlink property to navigate to an external web page.
// Common Searches: aspnet c# how to add a hyperlink label to an Excel chart with Aspose.Cells | Aspose.Cells set hyperlink on a shape displayed over a chart | C# add clickable text label to Excel chart using Aspose.Cells library | save Excel file with label shape that opens a web page when chart is opened Aspose.Cells
// Tags: Aspose.Cells add label hyperlink | C# Excel chart shape link | Aspose.Cells configure label screen tip | Excel workbook save with clickable label | Aspose.Cells format label blue underline

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example creates a new workbook, adds a column chart, inserts a label shape at the top-left corner, formats the label text as a blue underlined link, assigns a hyperlink to https://www.openai.com with a screen tip, and saves the file as ChartWithHyperlink.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a sample column chart (optional, just to have a chart on the sheet)
                int chartIdx = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = worksheet.Charts[chartIdx];
                // Populate chart data here if needed

                // Add a label shape that will act as a hyperlink
                // Parameters: upperLeftRow, upperLeftColumn, upperLeftRowOffset, upperLeftColumnOffset, height, width
                Label labelShape = worksheet.Shapes.AddLabel(0, 0, 0, 0, 30, 200);
                labelShape.Text = "Visit OpenAI";
                labelShape.Font.Color = Color.Blue;
                labelShape.Font.Underline = FontUnderlineType.Single;

                // Configure hyperlink for the label shape
                Hyperlink hyperlink = labelShape.Hyperlink;
                hyperlink.Address = "https://www.openai.com";
                hyperlink.ScreenTip = "OpenAI";

                // Define output file path
                string outputPath = "ChartWithHyperlink.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
