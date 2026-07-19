// Title: C# – Set chart area solid fill with 80 % opacity using Aspose.Cells
// Description: The example creates a workbook, adds sample data, inserts a column chart, and configures the chart area to a solid LightGray fill with 20 % transparency (80 % opacity) via FillFormat. The file is saved as ChartWithSolidFill.xlsx.
// Keywords: Aspose.Cells | C# | chart area fill | solid fill | opacity | transparency | FillFormat | column chart | Excel automation | chart styling
// Common Searches: Aspose.Cells set chart fill opacity | C# chart area solid color Aspose.Cells | how to make chart background semi‑transparent .NET | FillFormat transparency example | set chart background color Aspose.Cells
// Developer Intent: Apply a solid LightGray fill with 80 % opacity to a chart area.
// Use Cases: Design column charts with subtle background shading for professional reports. | Generate multiple charts in a workbook that share a consistent semi‑transparent fill to improve visual hierarchy. | Automate styling of chart elements in a reporting pipeline by adjusting FillFormat.Transparency.
// AI Prompts: Write C# code with Aspose.Cells that sets a chart area’s solid fill to LightGray at 80 % opacity. | Show how to use FillFormat.SolidFill.Color and FillFormat.SolidFill.Transparency to style a chart background in Aspose.Cells. | Explain the steps to apply a semi‑transparent solid fill to a chart area and save the workbook as an .xlsx file.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartFillExample
{
    // The example creates a workbook, adds sample data, inserts a column chart, and configures the chart area to a solid LightGray fill with 20 % transparency (80 % opacity) via FillFormat. The file is saved as ChartWithSolidFill.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure the chart area fill to a solid color with 80% opacity (20% transparency)
            chart.ChartArea.Area.FillFormat.FillType = FillType.Solid; // ensure solid fill
            chart.ChartArea.Area.FillFormat.SolidFill.Color = Color.LightGray; // subtle shading color
            chart.ChartArea.Area.FillFormat.SolidFill.Transparency = 0.2; // 0 = opaque, 1 = fully transparent

            // Save the workbook
            workbook.Save("ChartWithSolidFill.xlsx");
        }
    }
}
