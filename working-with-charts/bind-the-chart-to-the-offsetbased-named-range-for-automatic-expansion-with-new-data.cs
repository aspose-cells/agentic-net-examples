// Title: Create a dynamic column chart in C# by binding an Aspose.Cells chart series to an OFFSET‑based named range
// Common Searches: C# Aspose.Cells bind chart series to OFFSET named range for auto‑expanding data | How to use an OFFSET formula with a named range to make a dynamic chart in Aspose.Cells | Aspose.Cells chart automatically updates when new rows are added in C# | Define a dynamic data range with OFFSET in Aspose.Cells and link it to a column chart | Sample code for a dynamic Excel chart using an OFFSET named range with Aspose.Cells C#
// Tags: Aspose.Cells OFFSET named range | dynamic chart data range C# | column chart bound to named range | auto expanding Excel chart Aspose.Cells | C# Excel chart using OFFSET formula

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, defines an OFFSET‑based named range that grows as rows are added, inserts a column chart, binds the chart series to that named range, and saves the file as DynamicChart.xlsx.
class DynamicChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and set its name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Populate sample data (Date in column A, Value in column B)
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Value");
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(DateTime.Today.AddDays(i)); // Column A
                sheet.Cells[i + 1, 1].PutValue(i * 10 + 5);               // Column B
            }

            // Define an OFFSET‑based named range that expands automatically with new rows.
            // The range starts at A2 and includes both columns (Date and Value).
            string offsetFormula = "=OFFSET(Sheet1!$A$2,0,0,COUNTA(Sheet1!$A:$A)-1,2)";
            int nameIndex = workbook.Worksheets.Names.Add("DataRange");
            workbook.Worksheets.Names[nameIndex].RefersTo = offsetFormula;

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Dynamic Data Chart";

            // Bind the chart series to the OFFSET‑based named range (includes dates and values)
            chart.NSeries.Add("DataRange", true);

            // Save the workbook
            workbook.Save("DynamicChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
