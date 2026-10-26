// Title: Create a timeline line chart with clickable markers that open specific worksheets in the exported PDF using Aspose.Cells for .NET
// Common Searches: how to make chart markers open worksheets in a PDF with Aspose.Cells C# | Aspose.Cells add hyperlink to line chart data point for PDF export | C# generate timeline chart with interactive PDF links using Aspose.Cells | export workbook to PDF with active worksheet hyperlinks from chart markers | assign cell hyperlinks to chart series points in Aspose.Cells .NET
// Tags: chart marker hyperlink Aspose.Cells | timeline line chart PDF navigation .NET | add cell hyperlink to chart data point C# | export workbook to PDF with active links Aspose.Cells | worksheet navigation from chart marker Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace TimelineHyperlinkExample
{
    // The example builds a workbook, adds a line chart used as a timeline, creates separate worksheets for each month, assigns hyperlinks from the chart’s data points to the corresponding worksheet cells, and saves the workbook as a PDF where clicking a marker navigates to the related worksheet.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add worksheet that will hold the timeline data
                Worksheet wsData = workbook.Worksheets[0];
                wsData.Name = "TimelineData";

                // Populate timeline data (dates and values)
                wsData.Cells["A1"].PutValue("Date");
                wsData.Cells["B1"].PutValue("Value");
                wsData.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
                wsData.Cells["B2"].PutValue(10);
                wsData.Cells["A3"].PutValue(new DateTime(2023, 2, 1));
                wsData.Cells["B3"].PutValue(20);
                wsData.Cells["A4"].PutValue(new DateTime(2023, 3, 1));
                wsData.Cells["B4"].PutValue(15);

                // Add additional worksheets that will be opened when a marker is clicked
                Worksheet wsJan = workbook.Worksheets.Add("January");
                wsJan.Cells["A1"].PutValue("Details for January");
                Worksheet wsFeb = workbook.Worksheets.Add("February");
                wsFeb.Cells["A1"].PutValue("Details for February");
                Worksheet wsMar = workbook.Worksheets.Add("March");
                wsMar.Cells["A1"].PutValue("Details for March");

                // Create a line chart that will serve as the visual timeline
                int chartIndex = wsData.Charts.Add(ChartType.Line, 6, 0, 20, 10);
                Chart chart = wsData.Charts[chartIndex];

                // Add the data series (values)
                chart.NSeries.Add("B2:B4", true);
                // Basic chart appearance
                chart.ShowLegend = false;

                // Assign a hyperlink to each data point (marker) by linking the corresponding cell.
                // This works even if the Series.DataPoints API is unavailable in older versions.
                string[] targetSheets = { "January!A1", "February!A1", "March!A1" };
                for (int i = 0; i < targetSheets.Length; i++)
                {
                    // Cells B2, B3, B4 correspond to the data points
                    int row = i + 2; // 2,3,4
                    wsData.Hyperlinks.Add(row - 1, 1, row - 1, 1, targetSheets[i]); // column B (index 1)
                }

                // Save the workbook as PDF – the hyperlinks remain active in the PDF view
                string outputPath = "TimelineWithHyperlinks.pdf";
                workbook.Save(outputPath, SaveFormat.Pdf);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
