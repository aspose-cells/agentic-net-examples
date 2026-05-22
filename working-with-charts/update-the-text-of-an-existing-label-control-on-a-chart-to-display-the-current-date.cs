using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class UpdateChartLabel
{
    static void Main()
    {
        // Load an existing workbook that already contains a chart with a label
        Workbook workbook = new Workbook("ChartWithLabel.xlsx");

        // Get the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure there is at least one chart on the worksheet
        if (sheet.Charts.Count > 0)
        {
            // Access the first chart
            Chart chart = sheet.Charts[0];

            // ----- Update the chart title (a built‑in label) -----
            // Set the title text to the current date in desired format
            chart.Title.Text = DateTime.Now.ToString("yyyy-MM-dd");

            // ----- Update any custom label shapes added to the chart -----
            // Iterate through shapes that belong to the chart
            foreach (Shape shape in chart.Shapes)
            {
                // Identify label shapes (commonly TextBox auto shapes)
                if (shape.AutoShapeType == AutoShapeType.TextBox)
                {
                    // Set the shape's text to the current date
                    shape.Text = DateTime.Now.ToString("yyyy-MM-dd");
                }
            }
        }

        // Save the workbook with the updated label text
        workbook.Save("ChartWithLabel_Updated.xlsx");
    }
}