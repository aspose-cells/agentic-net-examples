// Title: Create a progress bar column chart with conditional formatting and threshold‑based colors using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set conditional formatting for progress values in Excel | How to color column chart points based on cell values with Aspose.Cells .NET | Apply three‑color threshold rules to an Excel chart series using Aspose.Cells | Programmatically add conditional formatting and chart point colors in C# Aspose.Cells | Create a progress bar chart with red, orange, and green colors using Aspose.Cells
// Tags: conditional formatting with Aspose.Cells C# | chart point color based on cell value Aspose.Cells | column chart series threshold colors Aspose.Cells | progress bar Excel chart Aspose.Cells | apply background colors to cells Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, writes task names and progress percentages, adds three conditional formatting rules (green < 30, orange 30‑69, red ≥ 70) to the progress cells, builds a column chart linked to the same data, and colors each chart point according to the same thresholds before saving the file as an XLSX workbook.
class Program
{
    static void Main()
    {
        const string outputPath = "ProgressBarConditionalFormatting.xlsx";

        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // -------------------------------------------------
            // 1. Populate sample progress values (0 – 100)
            // -------------------------------------------------
            string[] headers = { "Task", "Progress" };
            string[] tasks = { "Design", "Development", "Testing", "Deployment", "Review" };
            int[] progressValues = { 15, 45, 70, 85, 30 };

            // Write headers
            cells["A1"].PutValue(headers[0]);
            cells["B1"].PutValue(headers[1]);

            // Write data
            for (int i = 0; i < tasks.Length; i++)
            {
                cells[i + 2, 0].PutValue(tasks[i]);          // Column A
                cells[i + 2, 1].PutValue(progressValues[i]); // Column B
            }

            // -------------------------------------------------
            // 2. Apply Conditional Formatting to the Progress cells
            //    (Three rules: <30 = Green, 30‑69 = Orange, >=70 = Red)
            // -------------------------------------------------
            CellArea progressArea = new CellArea
            {
                StartRow = 1,               // B2 (zero‑based)
                EndRow = tasks.Length,      // B6
                StartColumn = 1,
                EndColumn = 1
            };

            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];
            cfCollection.AddArea(progressArea);

            // Rule 1: >= 70 → Red background
            int ruleRed = cfCollection.AddCondition(FormatConditionType.CellValue);
            FormatCondition fcRed = cfCollection[ruleRed];
            fcRed.Operator = OperatorType.GreaterOrEqual;
            fcRed.Formula1 = "70";
            fcRed.Style.BackgroundColor = Color.Red;
            fcRed.StopIfTrue = true;

            // Rule 2: >= 30 → Orange background
            int ruleOrange = cfCollection.AddCondition(FormatConditionType.CellValue);
            FormatCondition fcOrange = cfCollection[ruleOrange];
            fcOrange.Operator = OperatorType.GreaterOrEqual;
            fcOrange.Formula1 = "30";
            fcOrange.Style.BackgroundColor = Color.Orange;
            fcOrange.StopIfTrue = true;

            // Rule 3: < 30 → Green background
            int ruleGreen = cfCollection.AddCondition(FormatConditionType.CellValue);
            FormatCondition fcGreen = cfCollection[ruleGreen];
            fcGreen.Operator = OperatorType.LessThan;
            fcGreen.Formula1 = "30";
            fcGreen.Style.BackgroundColor = Color.LightGreen;

            // -------------------------------------------------
            // 3. Create a Column Chart that visualises the progress values
            // -------------------------------------------------
            int chartIndex = sheet.Charts.Add(ChartType.Column, 1, 3, 20, 12);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Task Progress";

            // Set the data range for the series (B2:B6)
            chart.NSeries.Add("B2:B6", true);
            // Set the category (X‑axis) labels (A2:A6)
            chart.NSeries.CategoryData = "A2:A6";

            // -------------------------------------------------
            // 4. Apply colors to each chart point based on the same thresholds
            // -------------------------------------------------
            Series series = chart.NSeries[0]; // Correct type is Series
            for (int i = 0; i < series.Points.Count; i++)
            {
                double val = Convert.ToDouble(cells[i + 1, 1].Value);

                Color pointColor = val >= 70
                    ? Color.Red
                    : val >= 30
                        ? Color.Orange
                        : Color.LightGreen;

                // Apply the color to the point's interior
                series.Points[i].Area.ForegroundColor = pointColor;
            }

            // -------------------------------------------------
            // 5. Save the workbook
            // -------------------------------------------------
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
