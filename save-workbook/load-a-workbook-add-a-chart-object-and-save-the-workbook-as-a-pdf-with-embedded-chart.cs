// Title: Create a column chart in an Excel worksheet and export the workbook to PDF with the chart embedded using Aspose.Cells for .NET (C#)
// AI Prompts: Generate a column chart from range A2:B3 in the first worksheet and save the workbook as a PDF file using Aspose.Cells in C#. | Add sample data, insert a column chart object, and convert the entire workbook to a PDF that includes the chart with Aspose.Cells .NET API.
// Common Searches: Aspose.Cells C# add column chart and convert workbook to PDF | how to embed an Excel chart into a PDF using Aspose.Cells .NET | programmatically create chart in worksheet and export to PDF Aspose.Cells | save workbook as PDF with chart included using C# Aspose.Cells | Aspose.Cells example converting Excel with chart to PDF
// Tags: Aspose.Cells create column chart C# | Aspose.Cells export workbook to PDF | Aspose.Cells embed chart in PDF | C# add chart to worksheet Aspose.Cells | convert Excel with chart to PDF Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartToPdfDemo
{
    // The code loads an existing Excel file, optionally writes sample data, creates a column chart on the first worksheet using the values in B2:B3 and categories in A2:A3, and then saves the workbook as a PDF where the chart is rendered inside the PDF.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // OPTIONAL: Add sample data for the chart if the workbook does not already contain it
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Fruits");
            worksheet.Cells["A3"].PutValue("Vegetables");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(50);
            worksheet.Cells["B3"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B3", true);          // Values
            chart.NSeries.CategoryData = "A2:A3";      // Categories

            // Save the entire workbook as a PDF; the chart will be embedded in the PDF output
            workbook.Save("output.pdf", SaveFormat.Pdf);

            Console.WriteLine("Workbook saved as PDF with embedded chart.");
        }
    }
}
