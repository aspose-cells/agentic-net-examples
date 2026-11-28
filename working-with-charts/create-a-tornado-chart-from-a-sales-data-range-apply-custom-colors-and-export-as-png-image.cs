// Title: Generate a tornado (horizontal bar) chart from sales data and export it as a PNG using Aspose.Cells for .NET (C#)
// Common Searches: c# generate tornado chart from Excel data using Aspose.Cells | how to set custom colors for series in an Aspose.Cells bar chart | export Aspose.Cells chart to PNG file in .NET | create sample sales workbook for tornado chart with Aspose.Cells | use negative values to create tornado effect in Aspose.Cells chart
// Tags: create tornado chart Aspose.Cells C# | custom series colors Aspose.Cells bar chart | png image export Aspose.Cells chart | horizontal bar chart for sales comparison Aspose.Cells | negative data values tornado effect Aspose.Cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace TornadoChartDemo
{
    // The example loads (or creates) a sales workbook, converts previous‑year values to negatives, adds a horizontal bar chart that functions as a tornado chart, assigns a title, defines category and two data series, applies distinct foreground colors to each series, and exports the chart directly to a PNG image while handling potential errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input and output files
            string inputPath = "sales_data.xlsx";
            string outputChartPath = "tornado_chart.png";

            try
            {
                // Load workbook if the input file exists; otherwise create a sample workbook
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                    Console.WriteLine($"Loaded workbook from \"{inputPath}\".");
                }
                else
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found. Creating sample data.");
                    workbook = CreateSampleWorkbook();
                }

                // Get the first worksheet (or create one if none)
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure the data range we will use exists
                PrepareDataForTornado(sheet);

                // Add a horizontal bar chart (which will serve as a tornado chart)
                int chartIndex = sheet.Charts.Add(ChartType.Bar, 10, 0, 30, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set chart title
                chart.Title.Text = "Sales Comparison (Tornado Chart)";

                // Define the data ranges
                chart.NSeries.CategoryData = "A2:A6";

                // Series 1 – Current Year (positive values)
                chart.NSeries.Add("B2:B6", false);
                chart.NSeries[0].Name = "Current Year";

                // Series 2 – Previous Year (negative values to appear on the opposite side)
                chart.NSeries.Add("C2:C6", false);
                chart.NSeries[1].Name = "Previous Year";

                // Apply custom colors to each series
                chart.NSeries[0].Area.ForegroundColor = Color.FromArgb(0, 112, 192); // Blue
                chart.NSeries[1].Area.ForegroundColor = Color.FromArgb(255, 192, 0); // Orange

                // Export the chart directly to a PNG file
                try
                {
                    chart.ToImage(outputChartPath, ImageType.Png);
                    Console.WriteLine($"Tornado chart successfully saved to \"{outputChartPath}\".");
                }
                catch (Exception imgEx)
                {
                    Console.WriteLine("Failed to export chart image:");
                    Console.WriteLine(imgEx.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the tornado chart:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }

        private static Workbook CreateSampleWorkbook()
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Header
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Current Year");
            ws.Cells["C1"].PutValue("Previous Year");

            // Sample categories
            string[] categories = { "Product A", "Product B", "Product C", "Product D", "Product E" };
            double[] currentYear = { 12000, 8500, 15000, 7000, 11000 };
            double[] previousYear = { 10000, 9000, 13000, 8000, 9500 };

            for (int i = 0; i < categories.Length; i++)
            {
                int row = i + 2; // Data starts at row 2
                ws.Cells[row, 0].PutValue(categories[i]);          // Column A
                ws.Cells[row, 1].PutValue(currentYear[i]);        // Column B (positive)
                ws.Cells[row, 2].PutValue(-previousYear[i]);      // Column C (negative for tornado effect)
            }

            // Save the sample workbook for reference (optional)
            string samplePath = "sample_sales_data.xlsx";
            wb.Save(samplePath, SaveFormat.Xlsx);
            Console.WriteLine($"Sample workbook saved to \"{samplePath}\".");

            return wb;
        }

        private static void PrepareDataForTornado(Worksheet sheet)
        {
            // Determine the used range for categories (assumes data starts at row 2)
            int firstDataRow = 2;
            int lastDataRow = sheet.Cells.MaxDataRow;

            // If there are no data rows, throw an informative exception
            if (lastDataRow < firstDataRow)
                throw new InvalidOperationException("The worksheet does not contain any data rows for the chart.");

            // Ensure column C (Previous Year) contains negative values
            for (int row = firstDataRow; row <= lastDataRow; row++)
            {
                Cell prevYearCell = sheet.Cells[row, 2]; // Column C (zero‑based index)
                if (prevYearCell.Type == CellValueType.IsNumeric)
                {
                    double val = prevYearCell.DoubleValue;
                    if (val > 0)
                        prevYearCell.PutValue(-val);
                }
                else
                {
                    // If the cell is empty or non‑numeric, set it to zero (neutral)
                    prevYearCell.PutValue(0);
                }
            }
        }
    }
}
