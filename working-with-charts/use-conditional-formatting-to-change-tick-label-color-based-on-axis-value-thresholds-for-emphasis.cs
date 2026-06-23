using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class ConditionalFormattingTickLabelColorDemo
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (categories in column A, values in column B)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Low");
            sheet.Cells["A3"].PutValue("Medium");
            sheet.Cells["A4"].PutValue("High");
            sheet.Cells["B2"].PutValue(15);   // Below low threshold
            sheet.Cells["B3"].PutValue(55);   // Between thresholds
            sheet.Cells["B4"].PutValue(95);   // Above high threshold

            // Add a column chart that uses the above data
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";      // Categories

            // Apply conditional formatting to the value cells (B2:B4)
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];

            // Define the range for conditional formatting
            CellArea cfArea = new CellArea { StartRow = 1, EndRow = 3, StartColumn = 1, EndColumn = 1 };
            cfCollection.AddArea(cfArea);

            // Add a 3‑color scale condition
            int conditionIdx = cfCollection.AddCondition(FormatConditionType.ColorScale);
            FormatCondition cf = cfCollection[conditionIdx];
            cf.ColorScale.Is3ColorScale = true;

            // Minimum – red (values < 30)
            cf.ColorScale.MinCfvo.Type = FormatConditionValueType.Number;
            cf.ColorScale.MinCfvo.Value = 0;
            cf.ColorScale.MinColor = Color.Red;

            // Midpoint – yellow (values between 30 and 70)
            cf.ColorScale.MidCfvo.Type = FormatConditionValueType.Number;
            cf.ColorScale.MidCfvo.Value = 50;
            cf.ColorScale.MidColor = Color.Yellow;

            // Maximum – green (values > 70)
            cf.ColorScale.MaxCfvo.Type = FormatConditionValueType.Number;
            cf.ColorScale.MaxCfvo.Value = 100;
            cf.ColorScale.MaxColor = Color.Green;

            // Adjust tick label appearance (font size, rotation, etc.)
            TickLabels tickLabels = chart.ValueAxis.TickLabels;
            tickLabels.Font.Size = 12;
            tickLabels.Font.IsBold = true;
            tickLabels.RotationAngle = 0; // horizontal

            // Save the workbook
            workbook.Save("ConditionalFormattingTickLabelColorDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}