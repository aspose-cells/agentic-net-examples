// Title: Aspose.Cells C# – Apply Built‑In Chart Style1 to Every Chart
// Description: Creates a workbook with sample data, adds two column charts, then loops through all worksheets and charts to set the Style property to 1 (Style1, range 1‑48). The workbook is saved as WorkbookWithStyledCharts.xlsx.
// Keywords: Aspose.Cells chart style C# | Style1 chart Aspose | apply built‑in chart style | set chart style programmatically | Aspose.Cells visual theme | C# Excel chart formatting | bulk chart styling Aspose
// Common Searches: how to set chart style1 in Aspose.Cells C# | apply built‑in style to all charts Aspose | chart style property range 1‑48 Aspose.Cells | iterate worksheets to change chart appearance | Aspose.Cells example for chart styling
// Developer Intent: Assign the predefined Style1 to each chart in a workbook using Aspose.Cells for C#.
// Use Cases: Enforce a consistent visual theme across automatically generated charts. | Prepare workbooks for corporate reporting where a specific chart style is required. | Refresh the look of existing charts without recreating them.
// AI Prompts: Generate C# code with Aspose.Cells that applies Style5 to all pie charts in a workbook. | List all built‑in chart style IDs and their descriptions using Aspose.Cells. | Show how to let a user select a chart style from a dropdown and apply it to every chart in a WinForms app.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook with sample data, adds two column charts, then loops through all worksheets and charts to set the Style property to 1 (Style1, range 1‑48). The workbook is saved as WorkbookWithStyledCharts.xlsx.
class ApplyChartStyle
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the charts
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B4"].PutValue(30);

        // Add first chart
        int chartIndex1 = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart1 = sheet.Charts[chartIndex1];
        chart1.NSeries.Add("B2:B4", false);
        chart1.NSeries.CategoryData = "A2:A4";

        // Add second chart
        int chartIndex2 = sheet.Charts.Add(ChartType.Column, 5, 10, 20, 18);
        Chart chart2 = sheet.Charts[chartIndex2];
        chart2.NSeries.Add("B2:B4", false);
        chart2.NSeries.CategoryData = "A2:A4";

        // Apply built‑in style 1 (Style1) to every chart in the workbook
        foreach (Worksheet ws in workbook.Worksheets)
        {
            foreach (Chart ch in ws.Charts)
            {
                ch.Style = 1; // Style1 (valid range: 1‑48)
            }
        }

        // Save the workbook
        workbook.Save("WorkbookWithStyledCharts.xlsx");
    }
}
