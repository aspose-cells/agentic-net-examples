// Title: How to add a column chart with cell‑based data labels on a hidden worksheet and then make the sheet visible using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to create a hidden worksheet, populate it with categories, values and label cells, insert a column chart that pulls its data labels from those cells, and finally change the worksheet visibility to visible before saving the workbook. | Demonstrate the steps to associate a chart series' data labels with a specific cell range on a concealed sheet and then reveal the worksheet using the Aspose.Cells API in a .NET console application.
// Common Searches: Aspose.Cells C# create chart on hidden sheet with data labels from cells | how to unhide a worksheet after adding a chart using Aspose.Cells .NET | link chart data labels to a cell range on a hidden worksheet Aspose.Cells | column chart with cell‑based labels on hidden worksheet Aspose.Cells example | set VisibilityType.Hidden then Visible after chart creation Aspose.Cells
// Tags: Aspose.Cells chart on concealed worksheet | cell‑based data labels for column chart | C# unhide worksheet after chart generation | associate chart series with cell range Aspose.Cells | VisibilityType.Hidden to Visible example

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartOnHiddenSheet
{
    // The sample creates a new workbook, adds a hidden worksheet named "HiddenData", fills columns A‑C with categories, numeric values, and label text, inserts a column chart that uses the values from column B and links its data labels to the text in column C, changes the worksheet's VisibilityType from Hidden to Visible, and saves the file as ChartOnHiddenWorksheet.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add a new worksheet that will be hidden initially
                Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenData");

                // Hide the worksheet
                hiddenSheet.VisibilityType = VisibilityType.Hidden;

                // Populate sample data
                // Categories
                hiddenSheet.Cells["A2"].PutValue("A");
                hiddenSheet.Cells["A3"].PutValue("B");
                hiddenSheet.Cells["A4"].PutValue("C");

                // Values for the chart
                hiddenSheet.Cells["B2"].PutValue(10);
                hiddenSheet.Cells["B3"].PutValue(20);
                hiddenSheet.Cells["B4"].PutValue(30);

                // Cell‑based labels (text to display on each point)
                hiddenSheet.Cells["C2"].PutValue("First");
                hiddenSheet.Cells["C3"].PutValue("Second");
                hiddenSheet.Cells["C4"].PutValue("Third");

                // Add a column chart on the hidden worksheet
                int chartIdx = hiddenSheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = hiddenSheet.Charts[chartIdx];

                // Set the data range for the series and categories
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable cell‑based data labels
                chart.NSeries[0].DataLabels.ShowCellRange = true;   // use cell range as labels
                chart.NSeries[0].DataLabels.LinkedSource = "C2:C4"; // link to label cells
                chart.NSeries[0].DataLabels.ShowValue = false;      // hide numeric values if desired

                // Make the worksheet visible again
                hiddenSheet.VisibilityType = VisibilityType.Visible;

                // Save the workbook
                string outputPath = "ChartOnHiddenWorksheet.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
