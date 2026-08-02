using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through each chart in the current worksheet
            foreach (Chart chart in sheet.Charts)
            {
                // Set the tick label direction of the category (X) axis to Rotate90
                chart.CategoryAxis.TickLabels.DirectionType = ChartTextDirectionType.Rotate90;

                // Set the tick label direction of the value (Y) axis to Rotate90
                chart.ValueAxis.TickLabels.DirectionType = ChartTextDirectionType.Rotate90;
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}