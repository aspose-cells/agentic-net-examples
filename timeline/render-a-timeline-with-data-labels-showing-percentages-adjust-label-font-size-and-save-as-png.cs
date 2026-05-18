using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate worksheet with sample data (dates, categories, amounts)
            cells["A1"].PutValue("Date");
            cells["B1"].PutValue("Category");
            cells["C1"].PutValue("Amount");

            // Define a style for date cells (m/d/yyyy)
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Number = 14; // Built‑in date format

            DateTime baseDate = new DateTime(2023, 1, 1);
            for (int i = 0; i < 5; i++)
            {
                cells[1 + i, 0].PutValue(baseDate.AddMonths(i)); // Date
                cells[1 + i, 0].SetStyle(dateStyle);            // Apply date format
                cells[1 + i, 1].PutValue("Item " + (i + 1));    // Category
                cells[1 + i, 2].PutValue((i + 1) * 10);         // Amount
            }

            // Create a pivot table based on the data
            PivotTableCollection pivots = sheet.PivotTables;
            int pivotIndex = pivots.Add("A1:C6", "E1", "PivotTable1");
            PivotTable pivot = pivots[pivotIndex];

            // Configure pivot fields
            pivot.AddFieldToArea(PivotFieldType.Page, "Date");      // Required for Timeline
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a Timeline control linked to the pivot table (Date field)
            sheet.Timelines.Add(pivot, 12, 0, "Date");

            // Add a pie chart that will display percentages in data labels
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 20, 0, 35, 15);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("C2:C6", true);
            chart.NSeries.CategoryData = "A2:A6";

            // Enable data labels and configure them to show percentages
            DataLabels dataLabels = chart.NSeries[0].DataLabels;
            dataLabels.ShowPercentage = true;   // display percentage values
            dataLabels.ShowValue = false;       // hide raw values
            dataLabels.Font.Size = 14;          // set desired font size
            dataLabels.ApplyFont();             // apply the font settings to all labels

            // Ensure output directory exists
            string outputPath = "TimelineChart.png";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Render the worksheet (including the timeline and chart) to a PNG image
            workbook.Save(outputPath, SaveFormat.Png);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}