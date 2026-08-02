// Title: Place Smart Marker Group Labels After Data Rows (LabelPosition='After') in Aspose.Cells for .NET
// Description: Learn how to configure a smart marker so that its group labels appear after the data rows by setting the LabelPosition attribute to "After". The example shows the XML smart‑marker definition, C# code to populate the workbook, and the resulting Excel file with correctly positioned labels.
// Keywords: Aspose.Cells smart markers C# | LabelPosition After Aspose.Cells | group label placement Excel | smart marker label position .NET | generate Excel with smart markers | C# set smart marker attributes
// Common Searches: how to set smart marker labelposition to after in Aspose.Cells | Aspose.Cells place group labels after rows | C# smart marker label position example | Aspose.Cells smart markers label after data | Excel smart marker group label placement
// Developer Intent: Add a smart marker to an Excel template that automatically positions its group labels after the generated data rows by using the LabelPosition='After' attribute.
// Use Cases: Creating financial statements where subtotal rows must appear below detail rows. | Generating inventory reports with category headers placed after item listings. | Automating invoice layouts that require totals to follow line‑item tables.
// AI Prompts: Generate C# code using Aspose.Cells that defines a smart marker with LabelPosition='After' to place group labels after data rows. | Show the XML smart‑marker syntax for setting LabelPosition to After and explain how Aspose.Cells interprets it. | Explain step‑by‑step how to bind data to a smart marker template so that group labels are rendered after the data rows in the final workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // Learn how to configure a smart marker so that its group labels appear after the data rows by setting the LabelPosition attribute to "After". The example shows the XML smart‑marker definition, C# code to populate the workbook, and the resulting Excel file with correctly positioned labels.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Create a column chart and set its position
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];

                // Add series data (values) and category (X) data
                chart.NSeries.Add("B2:B4", true);
                // Use XValues to specify category data (compatible with current Aspose.Cells version)
                chart.NSeries[0].XValues = "A2:A4";

                // Optional: set chart title
                chart.Title.Text = "Sample Chart";

                // Save the resulting workbook
                string outputPath = "SmartMarkerChart_With_AfterLabel.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
