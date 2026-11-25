// Title: Link a column chart series to a named range for dynamic updates using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# link chart series to a named range for dynamic data | how to create a named range and bind it to a column chart in .NET | dynamic Excel chart that updates when a named range changes using Aspose.Cells | C# example of adding a named range and using it as a chart data source
// Tags: chart series binding to named range Aspose.Cells | dynamic chart data source Excel .NET | create named range for chart series C# | column chart linked to named range Aspose.Cells | Aspose.Cells chart series data binding example

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartNamedRangeDemo
{
    // The sample creates a workbook, defines a named range "SalesData" covering cells B2:B6, adds a column chart, and links the chart's series to the named range so the chart updates automatically when the range changes, then saves the file as ChartWithNamedRange.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data
                cells["A1"].PutValue("Product");
                cells["B1"].PutValue("Sales");

                string[] products = { "Apple", "Banana", "Cherry", "Date", "Elderberry" };
                int[] sales = { 120, 150, 90, 200, 130 };

                for (int i = 0; i < products.Length; i++)
                {
                    // Column A – Product names
                    cells[i + 1, 0].PutValue(products[i]);
                    // Column B – Sales figures
                    cells[i + 1, 1].PutValue(sales[i]);
                }

                // Define a named range "SalesData" that refers to the sales column (B2:B6)
                int nameIndex = workbook.Worksheets.Names.Add("SalesData");
                Name salesName = workbook.Worksheets.Names[nameIndex];
                string startCell = $"${CellsHelper.ColumnIndexToName(1)}${2}"; // B2
                string endCell = $"${CellsHelper.ColumnIndexToName(1)}${products.Length + 1}"; // B6
                salesName.RefersTo = $"=Sheet1!{startCell}:{endCell}";

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];
                chart.Title.Text = "Dynamic Sales Chart";

                // Link the chart series to the named range "SalesData"
                chart.NSeries.Add("=SalesData", true);

                // Optionally set category (X‑axis) data – using the product names range
                // This line is omitted because the Series class in the current Aspose.Cells version
                // does not expose a CategoryData property. The chart will use the default categories.

                // Save the workbook
                string outputPath = "ChartWithNamedRange.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("The chart series is linked to the named range 'SalesData' and will update dynamically when the range changes.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
