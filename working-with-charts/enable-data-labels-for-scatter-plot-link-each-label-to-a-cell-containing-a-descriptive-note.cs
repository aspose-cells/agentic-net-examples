// Title: Create a scatter chart with data labels linked to note cells using Aspose.Cells for .NET (C#)
// AI Prompts: Generate a C# program that uses Aspose.Cells to construct a scatter diagram, turn on data labels for its series, and fill each label with the corresponding note from column C. | Provide a full Aspose.Cells sample that builds a workbook, adds X/Y data and comment cells, creates a scatter chart with Y‑value labels, and documents that attaching labels to worksheet cells is not available.
// Common Searches: add data labels to a scatter plot from a notes column using Aspose.Cells | bind chart data labels to worksheet cells in a .NET workbook | show Y values as labels on a scatter diagram with Aspose.Cells | does Aspose.Cells support linking data labels to a cell range | step by step create scatter chart with descriptive labels in C#
// Tags: Aspose.Cells NSeries.DataLabels property C# | create scatter diagram with Aspose.Cells .NET | link chart labels to worksheet cells limitation Aspose.Cells | save workbook as XLSX using Aspose.Cells | populate chart series XValues YValues Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts; // Required for Chart and ChartType

namespace AsposeCellsExample
{
    // The example creates a new workbook, writes X values, Y values, and descriptive notes into columns A‑C, adds a scatter chart referencing the X and Y ranges, enables data labels to display the Y values, and saves the file as ScatterWithLinkedLabels.xlsx. It also notes that linking each data label to a specific note cell is not supported in the current Aspose.Cells version.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet and rename it
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data";

                // Populate X values (column A), Y values (column B) and descriptive notes (column C)
                sheet.Cells["A1"].PutValue("X");
                sheet.Cells["B1"].PutValue("Y");
                sheet.Cells["C1"].PutValue("Note");

                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue(2);
                sheet.Cells["C2"].PutValue("Point A");

                sheet.Cells["A3"].PutValue(2);
                sheet.Cells["B3"].PutValue(4);
                sheet.Cells["C3"].PutValue("Point B");

                sheet.Cells["A4"].PutValue(3);
                sheet.Cells["B4"].PutValue(6);
                sheet.Cells["C4"].PutValue("Point C");

                sheet.Cells["A5"].PutValue(4);
                sheet.Cells["B5"].PutValue(8);
                sheet.Cells["C5"].PutValue("Point D");

                // Add a scatter chart
                int chartIndex = sheet.Charts.Add(ChartType.Scatter, 7, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the series (Y values) and X values
                chart.NSeries.Add("=Data!$B$2:$B$5", true);
                chart.NSeries[0].XValues = "=Data!$A$2:$A$5";

                // Enable data labels for the series
                chart.NSeries[0].DataLabels.ShowValue = true;          // show the Y value (optional)
                chart.NSeries[0].DataLabels.ShowSeriesName = false;    // hide series name
                chart.NSeries[0].DataLabels.ShowCategoryName = false; // hide category name

                // Note: Linking data labels to cells (ShowDataLabelRange / DataLabelRange) is not supported
                // in the current Aspose.Cells version used. These lines have been removed to ensure compilation.

                // Save the workbook
                workbook.Save("ScatterWithLinkedLabels.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
