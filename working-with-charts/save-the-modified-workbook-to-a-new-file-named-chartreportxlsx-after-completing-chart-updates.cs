// Title: Create and Save an Excel Workbook with a Column Chart (ChartReport.xlsx) using Aspose.Cells for .NET (C#)
// Description: A C# sample that builds a new workbook, fills cells A1:B4 with category and value data, inserts a column chart, assigns a title, and writes the file as ChartReport.xlsx via Aspose.Cells for .NET.
// Keywords: Aspose.Cells | .NET | C# | column chart | Excel chart example | save workbook | ChartReport.xlsx | create Excel file | chart title | NSeries | ChartType.Column
// Common Searches: Aspose.Cells add column chart C# | save workbook with specific filename Aspose.Cells | C# example create Excel chart and export | set chart title Aspose.Cells | generate chart report using Aspose.Cells
// Developer Intent: Produce an Excel file that contains a column chart and persist it as ChartReport.xlsx.
// Use Cases: Automate a sales summary that visualizes product categories with a column chart for email distribution. | Create a monthly KPI dashboard where data is populated programmatically, a chart is added, and the workbook is saved for executive review. | Export query results from a database to Excel and include a column chart for quick visual analysis.
// AI Prompts: Provide C# code that adds a line chart to the same workbook and saves it as LineReport.xlsx using Aspose.Cells. | Explain how to modify the column chart style, move the legend, and enable data labels in the given example. | Show how to load an existing template workbook, replace its data range, refresh the chart, and save the updated file.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ChartReportGenerator
{
    // A C# sample that builds a new workbook, fills cells A1:B4 with category and value data, inserts a column chart, assigns a title, and writes the file as ChartReport.xlsx via Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
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

            // Perform any additional chart updates here (e.g., title, style)
            chart.Title.Text = "Sample Column Chart";

            // Save the modified workbook to a new file named ChartReport.xlsx
            workbook.Save("ChartReport.xlsx", SaveFormat.Xlsx);
        }
    }
}
