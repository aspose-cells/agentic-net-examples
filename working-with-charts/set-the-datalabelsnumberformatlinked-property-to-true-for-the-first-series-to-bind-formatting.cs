// Title: Aspose.Cells C# – Set DataLabels.NumberFormatLinked = true for the first chart series
// Description: Creates a workbook, adds sample data, inserts a column chart, links the first series' data labels to cells C2:C3, shows label values, and enables NumberFormatLinked so the labels inherit the number format from the source cells before saving the file.
// Keywords: Aspose.Cells | C# chart data labels | NumberFormatLinked | LinkedSource | chart series formatting | Excel chart automation | Aspose.Cells example
// Common Searches: Aspose.Cells set NumberFormatLinked C# | link data label format to source cells Aspose.Cells | chart series data label formatting .NET | how to bind number format of data labels Aspose.Cells | Aspose.Cells column chart data labels example
// Developer Intent: Enable the NumberFormatLinked property for the first series' data labels so they automatically use the number format defined in the linked source cells.
// Use Cases: Generate a column chart where data labels display custom‑formatted values (e.g., "100 units") defined in worksheet cells. | Build a reporting workbook that keeps label styling in sync with source cell formats, reducing manual formatting effort. | Create dynamic Excel charts that automatically reflect changes to number formats in the linked source range.
// AI Prompts: Show how to set DataLabels.NumberFormatLinked = true for a chart series using Aspose.Cells in C# and describe its impact. | Provide a step‑by‑step C# example that creates a chart, links data label values to a range, and binds the number format to the source cells. | Generate code that updates chart data label formatting after modifying the number format of the linked source cells with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds sample data, inserts a column chart, links the first series' data labels to cells C2:C3, shows label values, and enables NumberFormatLinked so the labels inherit the number format from the source cells before saving the file.
    public class SetNumberFormatLinkedDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(100);
                worksheet.Cells["B3"].PutValue(200);
                worksheet.Cells["C1"].PutValue("Formatted Value");
                worksheet.Cells["C2"].PutValue("100 units");
                worksheet.Cells["C3"].PutValue("200 units");

                // Add a column chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = worksheet.Charts[chartIndex];

                // Define the series data range and category labels
                chart.NSeries.Add("B2:B3", true);
                chart.NSeries.CategoryData = "A2:A3";

                // Access the first series
                Series series = chart.NSeries[0];

                // Enable data labels and bind them to the source cells
                series.DataLabels.ShowValue = true;
                series.DataLabels.LinkedSource = "C2:C3";

                // Bind the number format of the data labels to the source cells
                series.DataLabels.NumberFormatLinked = true;

                // Save the workbook to an XLSX file
                workbook.Save("SetNumberFormatLinkedDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetNumberFormatLinkedDemo.Run();
        }
    }
}
