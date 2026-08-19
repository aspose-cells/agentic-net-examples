// Title: C# – Set Chart Axis Tick Labels to Horizontal with Aspose.Cells
// Description: Load an Excel workbook, locate a chart, and set the CategoryAxis and ValueAxis tick‑label direction to horizontal using Aspose.Cells for .NET, then save the file.
// Keywords: Aspose.Cells C# chart axis tick labels | horizontal tick label direction | ChartTextDirectionType.Horizontal | set category axis label orientation | modify Excel chart programmatically | Aspose.Cells chart example | C# Excel chart label direction
// Common Searches: Aspose.Cells set chart tick label horizontal | C# change Excel chart axis label orientation | ChartTextDirectionType.Horizontal Aspose.Cells | how to make chart axis labels horizontal in .NET | set category axis tick labels direction Aspose.Cells
// Developer Intent: Change the orientation of tick labels on a chart’s axes to horizontal in an Excel file using Aspose.Cells for .NET.
// Use Cases: Generate financial dashboards where X‑axis dates must stay horizontal for clarity. | Prepare printable reports with consistent label orientation across multiple charts. | Automate chart styling in batch processing of Excel workbooks.
// AI Prompts: Generate C# code that iterates through all worksheets and sets every chart’s CategoryAxis and ValueAxis tick label direction to horizontal using Aspose.Cells. | Show how to set tick label direction to vertical or rotated for a specific axis in Aspose.Cells. | Explain how to detect if a chart has a CategoryAxis before applying ChartTextDirectionType.Horizontal in C#.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Load an Excel workbook, locate a chart, and set the CategoryAxis and ValueAxis tick‑label direction to horizontal using Aspose.Cells for .NET, then save the file.
class SetTickLabelsDirection
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (or specify the appropriate one)
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure the worksheet contains at least one chart
        if (worksheet.Charts.Count == 0)
        {
            Console.WriteLine("No charts found in the worksheet.");
            return;
        }

        // Get the first chart in the worksheet
        Chart chart = worksheet.Charts[0];

        // Set tick labels direction to horizontal for the category (X) axis
        chart.CategoryAxis.TickLabels.DirectionType = ChartTextDirectionType.Horizontal;

        // Optionally, also set the value (Y) axis tick labels direction to horizontal
        chart.ValueAxis.TickLabels.DirectionType = ChartTextDirectionType.Horizontal;

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
