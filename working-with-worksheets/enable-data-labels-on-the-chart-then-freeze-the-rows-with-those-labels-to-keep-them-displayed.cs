// Title: Enable data labels on a column chart and freeze the top rows in an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells show values and category names on column chart data labels | How to freeze the first five rows and a column in an Excel sheet using Aspose.Cells | Aspose.Cells set data label position to inside end for a column chart series | Example of adding data labels and applying freeze panes in the same workbook with Aspose.Cells | Save an Excel file with frozen header rows and customized chart labels using Aspose.Cells .NET
// Tags: add data labels to column chart Aspose.Cells C# | customize chart data label font and background Aspose.Cells | set data label position inside end Aspose.Cells | freeze panes for header rows Aspose.Cells | save workbook with frozen panes Xlsx Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace EnableDataLabelsAndFreezeRows
{
    // The example creates a workbook, fills it with sample data, adds a column chart, enables data labels that display both values and category names with custom font and background, positions the labels inside the columns, freezes the first five rows and the first column, and saves the file as an XLSX workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            // Header row
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            // Data rows
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(90);
            sheet.Cells["A5"].PutValue("D");
            sheet.Cells["B5"].PutValue(200);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B5", true);          // Values
            chart.NSeries.CategoryData = "A2:A5";      // Categories

            // Enable data labels for the first series
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;        // Show the numeric value
            series.DataLabels.ShowCategoryName = true; // Show the category name
            series.DataLabels.Position = LabelPositionType.InsideEnd; // Position inside the column

            // Optional: customize appearance of data labels
            series.DataLabels.Font.IsBold = true;
            series.DataLabels.Font.Color = System.Drawing.Color.DarkBlue;
            series.DataLabels.BackgroundMode = BackgroundMode.Transparent;

            // Freeze the rows that contain the data labels (rows 1‑5)
            // This will keep the header and data rows visible while scrolling
            // FreezePanes(row, column, freezedRows, freezedColumns)
            // Freeze first 5 rows and first column (A) as the left pane
            sheet.FreezePanes(5, 1, 5, 1);

            // Save the workbook to a file
            string outputPath = "ChartWithDataLabelsAndFrozenRows.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to '{outputPath}'.");
            Console.WriteLine("Data labels are enabled on the chart, and rows 1‑5 are frozen.");
        }
    }
}
