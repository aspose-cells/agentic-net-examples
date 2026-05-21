using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Data";

        // Populate initial data (header + 5 rows)
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        for (int i = 2; i <= 6; i++)
        {
            sheet.Cells[i - 1, 0].PutValue("Item " + (i - 1));
            sheet.Cells[i - 1, 1].PutValue(i * 10);
        }

        // Create an OFFSET‑based named range that expands automatically
        // Starts at A2, height = number of non‑empty cells in column A minus the header,
        // width = 2 columns (A and B)
        int nameIndex = workbook.Worksheets.Names.Add("DynamicRange");
        Name dynamicRange = workbook.Worksheets.Names[nameIndex];
        dynamicRange.SetRefersTo("=OFFSET(Data!$A$2,0,0,COUNTA(Data!$A:$A)-1,2)", false, false);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 8, 0, 25, 7);
        Chart chart = sheet.Charts[chartIndex];

        // Bind the chart to the dynamic named range
        chart.NSeries.Add("=DynamicRange", true);

        // Optional: set chart title
        chart.Title.Text = "Dynamic Data Chart";

        // Save the workbook
        workbook.Save("DynamicChart.xlsx");
    }
}