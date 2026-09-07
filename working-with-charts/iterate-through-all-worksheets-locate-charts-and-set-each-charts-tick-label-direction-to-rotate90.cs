// Title: Rotate X and Y axis tick labels 90° for every chart across all worksheets with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook using Aspose.Cells, loops through each worksheet and every chart, and sets both the category and value axis tick label DirectionType to Rotate90 before saving the file. | Write a script that employs Aspose.Cells for .NET to change the orientation of all chart axis tick labels to 90 degrees in every sheet of a given workbook.
// Common Searches: Aspose.Cells C# rotate chart axis labels to vertical | How to set tick label direction for all charts in an Excel workbook using Aspose.Cells | Programmatically change Excel chart axis label orientation with Aspose.Cells .NET | Iterate through worksheets and modify chart properties Aspose.Cells
// Tags: set chart axis tick label direction Aspose.Cells | rotate chart tick labels 90 degrees .NET | iterate worksheets modify charts Aspose.Cells | categoryaxis ticklabels directiontype Rotate90 | valueaxis ticklabels directiontype Rotate90

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads a workbook, iterates over each worksheet and its charts, sets the CategoryAxis and ValueAxis tick label DirectionType to Rotate90, and saves the updated workbook.
class SetChartTickLabelDirection
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all charts on the current worksheet
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
