// Title: Use Aspose.Cells for .NET to define a dynamic named range that auto‑expands and link it to a column chart
// AI Prompts: Generate C# code with Aspose.Cells that creates a named range using OFFSET and COUNTA to automatically include new rows. | Show how to attach the dynamic named range to a column chart and update the chart after adding additional data rows. | Include steps to recalculate workbook formulas and save the workbook after the range expands.
// Common Searches: Aspose.Cells C# create dynamic named range that grows with added rows | how to bind a dynamic named range to a chart in Aspose.Cells .NET | refresh chart data source after appending rows using Aspose.Cells | recalculate formulas for OFFSET named range in Aspose.Cells workbook | example of using OFFSET and COUNTA for auto‑expanding chart range in C#
// Tags: dynamic named range OFFSET COUNTA Aspose.Cells | auto expanding chart source C# | recalculate workbook formulas Aspose.Cells | bind named range to column chart .NET | dynamic chart data range Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// // Demonstrates creating a workbook, defining a dynamic named range with OFFSET/COUNTA, adding a column chart that references the range, appending more rows, recalculating formulas, and saving the file.
class DynamicNamedRangeDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // Fill initial data in column A (A1:A5)
            for (int i = 0; i < 5; i++)
            {
                cells[i, 0].PutValue(i + 1);
            }

            // Define a dynamic named range "MyData" that expands with non‑empty cells in column A
            int nameIdx = wb.Worksheets.Names.Add("MyData");
            Name dynName = wb.Worksheets.Names[nameIdx];
            dynName.RefersTo = "=OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)";

            // Create a column chart that uses the dynamic named range as its data source
            int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            Chart chart = ws.Charts[chartIdx];
            chart.NSeries.Add("MyData", true);
            chart.Title.Text = "Dynamic Data";

            // Append more data to column A after the chart is created
            for (int i = 5; i < 10; i++)
            {
                cells[i, 0].PutValue(i + 1);
            }

            // Recalculate formulas so the dynamic named range reflects the new rows
            wb.CalculateFormula();

            // Save the workbook
            string outputPath = "DynamicNamedRangeDemo.xlsx";
            wb.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
