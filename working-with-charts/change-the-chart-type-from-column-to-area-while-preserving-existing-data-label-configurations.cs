// Title: Convert a Column Chart to an Area Chart in C# with Aspose.Cells while Keeping Data Labels
// Description: A concise C# example that loads an Excel workbook, selects the first chart, switches its type from Column to Area using Aspose.Cells, and saves the file. The operation retains all existing data‑label settings and series formatting.
// Keywords: Aspose.Cells chart type change C# | convert column chart to area chart .NET | preserve data labels Aspose.Cells | C# change Excel chart type programmatically | Aspose.Cells Area chart example | Excel chart conversion Aspose | global
// Common Searches: how to change a column chart to an area chart using Aspose.Cells | Aspose.Cells keep data labels when changing chart type | C# programmatically convert Excel chart type | Aspose.Cells chart type Area example | preserve series formatting Aspose.Cells chart conversion
// Developer Intent: Switch an existing column chart to an area chart without resetting data‑label or series properties.
// Use Cases: Update legacy financial dashboards to area charts while retaining label formatting. | Batch‑process multiple workbooks to replace column charts with area charts automatically. | Build a user‑driven tool that lets end users toggle chart types without losing custom label settings.
// AI Prompts: Write C# code with Aspose.Cells that changes any Column chart to an Area chart and ensures data labels stay unchanged. | Show how to loop through all charts in a worksheet, detect Column charts, and convert them to Area charts while preserving series styles. | Explain Aspose.Cells' behavior for data‑label preservation when Chart.Type is modified and what extra steps are required for custom label positions.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartTypeChange
{
    // A concise C# example that loads an Excel workbook, selects the first chart, switches its type from Column to Area using Aspose.Cells, and saves the file. The operation retains all existing data‑label settings and series formatting.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook that already contains a column chart
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Assume the chart we want to modify is the first chart in the collection
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            Chart chart = sheet.Charts[0];

            // Preserve existing data label settings – no action needed because
            // changing the chart type does not reset series or their data labels.
            // Change the chart type from Column to Area
            chart.Type = ChartType.Area;

            // Save the modified workbook
            workbook.Save("output.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Chart type changed to Area and workbook saved as output.xlsx");
        }
    }
}
