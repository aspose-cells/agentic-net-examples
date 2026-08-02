using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDynamicChart
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate sample data (A column = categories, B column = values)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            for (int i = 2; i <= 11; i++)
            {
                sheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
                sheet.Cells[$"B{i}"].PutValue(i * 10);
            }

            // Define a dynamic named range using OFFSET.
            // The range will start at B2 and expand vertically based on the number of non‑empty cells in column B.
            // OFFSET(Sheet1!$B$2,0,0,COUNTA(Sheet1!$B:$B)-1,1)
            int nameIndex = workbook.Worksheets.Names.Add("DynamicValues");
            Name dynamicName = workbook.Worksheets.Names[nameIndex];
            // Set the RefersTo formula (false = A1 style, false = not locale‑specific)
            dynamicName.SetRefersTo("=OFFSET(Data!$B$2,0,0,COUNTA(Data!$B:$B)-1,1)", false, false);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 13, 0, 30, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Bind the chart to the dynamic named range.
            // Using NSeries.Add with the named range reference.
            chart.NSeries.Add("=DynamicValues", true);
            chart.Title.Text = "Dynamic Data Chart";

            // Save the workbook
            workbook.Save("DynamicChart.xlsx");
        }
    }
}