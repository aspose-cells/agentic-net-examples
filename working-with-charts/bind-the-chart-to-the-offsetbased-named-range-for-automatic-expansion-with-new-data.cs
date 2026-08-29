// Title: Create a column chart bound to an OFFSET‑based dynamic named range using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# with Aspose.Cells that defines a self‑adjusting named range using an OFFSET formula and links it to a column chart series. | Show how to embed a COUNTA function inside an OFFSET expression to calculate the height of an auto‑growing range for chart data in Aspose.Cells. | Provide the steps to add a column chart to a worksheet, bind its series to the defined range, and save the workbook with Aspose.Cells.
// Common Searches: asp.net aspose.cells bind chart series to offset named range | c# create auto expanding chart data range with counTA and offset in Aspose.Cells | how to make a column chart update automatically when new rows are added in Aspose.Cells | using offset formula for self adjusting range in Aspose.Cells chart | asp.net aspose.cells chart series from expanding named range
// Tags: OFFSET dynamic named range Aspose.Cells | auto expanding chart source range C# | column chart series from named range Aspose.Cells | COUNTA in OFFSET formula Aspose.Cells | Aspose.Cells chart binding dynamic range .NET

using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, defines an OFFSET‑based named range that automatically expands with new rows using COUNTA, adds a column chart, binds the chart series to the dynamic range, and saves the file as DynamicChart.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        ws.Name = "Sheet1";

        // Populate initial data (header + 5 rows)
        ws.Cells["A1"].PutValue("Category");
        ws.Cells["B1"].PutValue("Value");
        for (int i = 1; i <= 5; i++)
        {
            ws.Cells[i + 1, 0].PutValue("Item " + i);   // Column A
            ws.Cells[i + 1, 1].PutValue(i * 10);        // Column B
        }

        // Define an OFFSET‑based named range that expands automatically with new rows
        // Starts at A2 (first data row), height = number of non‑empty rows in column A minus the header, width = 2 columns
        int nameIdx = wb.Worksheets.Names.Add("DynamicData");
        Name dynName = wb.Worksheets.Names[nameIdx];
        dynName.RefersTo = "=OFFSET(Sheet1!$A$2,0,0,COUNTA(Sheet1!$A:$A)-1,2)";

        // Add a column chart
        int chartIdx = ws.Charts.Add(ChartType.Column, 7, 0, 25, 7);
        Chart chart = ws.Charts[chartIdx];

        // Bind the chart to the dynamic named range
        // The series formula references the named range; 'true' indicates data are organized by column
        chart.NSeries.Add("=DynamicData", true);

        // Optional: set chart title
        chart.Title.Text = "Dynamic Data Chart";

        // Save the workbook
        wb.Save("DynamicChart.xlsx");
    }
}
