using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (contains a default sheet)
            Workbook workbook = new Workbook();

            // Remove the default sheet to avoid duplicate names
            if (workbook.Worksheets.Count > 0)
                workbook.Worksheets.RemoveAt(0);

            // Add several worksheets with unique names
            workbook.Worksheets.Add("Sheet1");
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // Loop through each worksheet to add data and a chart
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];

                // Populate sample data (Category / Value)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                for (int r = 2; r <= 6; r++)
                {
                    sheet.Cells[$"A{r}"].PutValue($"Item {r - 1}");
                    sheet.Cells[$"B{r}"].PutValue((r - 1) * 10);
                }

                // Add a column chart to the current worksheet
                int chartIdx = sheet.Charts.Add(ChartType.Column, 8, 0, 20, 8);
                Chart chart = sheet.Charts[chartIdx];

                // Define the data range for the chart
                chart.NSeries.Add("B2:B6", true);          // Values
                chart.NSeries.CategoryData = "A2:A6";      // Categories

                // Set a title for clarity
                chart.Title.Text = $"Chart for {sheet.Name}";

                // Recalculate the chart layout before saving
                chart.Calculate();
            }

            // Save the workbook with all charts
            workbook.Save("ChartsForAllSheets.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}