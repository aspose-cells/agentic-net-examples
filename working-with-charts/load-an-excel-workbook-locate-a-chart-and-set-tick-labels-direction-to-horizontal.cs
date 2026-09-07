// Title: Set category axis tick labels to horizontal in an Excel chart using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an existing .xlsx workbook with Aspose.Cells, retrieves the first chart, and changes the CategoryAxis tick labels direction to horizontal before saving the file. | Show how to use the Aspose.Cells Chart API in a .NET project to modify the text direction of a chart's axis tick labels to horizontal.
// Common Searches: Aspose.Cells C# set chart X axis tick label direction horizontal | change category axis tick label orientation to horizontal in Excel using Aspose.Cells | C# Aspose.Cells modify chart axis text direction horizontal
// Tags: Aspose.Cells chart axis tick label orientation | C# set category axis text direction | horizontal tick labels Excel chart | modify chart axis label direction Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing workbook (input.xlsx), accesses the first worksheet and its first chart, sets the CategoryAxis tick labels direction to horizontal, and saves the updated workbook as output.xlsx using Aspose.Cells for .NET.
class SetTickLabelsDirection
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure the worksheet contains at least one chart
        if (worksheet.Charts.Count > 0)
        {
            // Get the first chart in the worksheet
            Chart chart = worksheet.Charts[0];

            // Access the category (X) axis tick labels
            TickLabels tickLabels = chart.CategoryAxis.TickLabels;

            // Set the tick labels direction to horizontal
            tickLabels.DirectionType = ChartTextDirectionType.Horizontal;
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
