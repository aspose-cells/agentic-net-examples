// Title: Create pie charts for every data table in all worksheets of an Excel workbook using Aspose.Cells for .NET
// Common Searches: C# generate a pie chart for each worksheet in an existing Excel file with Aspose.Cells | how to loop through worksheets and insert pie charts programmatically using Aspose.Cells .NET | batch add pie charts to multiple sheets based on A2:B range in Aspose.Cells | automate pie chart creation for all tables in an Excel workbook using C# | Aspose.Cells place pie chart below data table for each sheet in a workbook
// Tags: Aspose.Cells insert pie chart on worksheet | C# loop through worksheets Aspose.Cells | define chart data range A1 notation Aspose.Cells | batch generate Excel charts with .NET | automate chart placement below data table Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace BatchPieChartGenerator
{
    // The example loads an existing workbook (or creates a sample one), iterates over every worksheet, checks for sufficient data, defines the range A2:B‑lastRow, adds a pie chart positioned below the table, sets the chart title and data labels, and saves the workbook with a pie chart on each sheet.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input and output workbooks
            string inputPath = "input.xlsx";
            string outputPath = "output_with_pie_charts.xlsx";

            try
            {
                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a sample workbook
                if (File.Exists(inputPath))
                {
                    Console.WriteLine($"Loading workbook from \"{inputPath}\"...");
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a sample workbook.");
                    workbook = CreateSampleWorkbook();
                }

                // Process each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    try
                    {
                        Console.WriteLine($"Processing worksheet \"{sheet.Name}\"...");

                        // Determine the used range
                        Cells cells = sheet.Cells;
                        int maxRow = cells.MaxDataRow;      // zero‑based index of last row with data
                        int maxColumn = cells.MaxDataColumn; // zero‑based index of last column with data

                        // Require at least one data row (header + data) and at least two columns
                        if (maxRow < 1 || maxColumn < 1)
                        {
                            Console.WriteLine($"  Skipping \"{sheet.Name}\" – not enough data for a pie chart.");
                            continue;
                        }

                        // Define data range (A1‑style)
                        int dataStartRow = 2; // first data row (1‑based, after header)
                        int dataEndRow = maxRow + 1; // convert zero‑based to 1‑based
                        string dataRange = $"A{dataStartRow}:B{dataEndRow}";

                        // Add a pie chart below the data table
                        int chartTopRow = maxRow + 3; // one empty row after data
                        int chartLeftColumn = 0;      // column A
                        int chartBottomRow = chartTopRow + 15;
                        int chartRightColumn = chartLeftColumn + 7;

                        int chartIndex = sheet.Charts.Add(ChartType.Pie, chartTopRow, chartLeftColumn, chartBottomRow, chartRightColumn);
                        Chart pieChart = sheet.Charts[chartIndex];

                        // Set the data source for the chart
                        pieChart.SetChartDataRange(dataRange, true);

                        // Set chart title
                        pieChart.Title.Text = $"Pie Chart – {sheet.Name}";
                        pieChart.Title.IsVisible = true;

                        // Show data labels (value)
                        if (pieChart.NSeries.Count > 0)
                        {
                            pieChart.NSeries[0].DataLabels.ShowValue = true;
                        }

                        Console.WriteLine($"  Pie chart created for \"{sheet.Name}\" using range {dataRange}.");
                    }
                    catch (Exception wsEx)
                    {
                        Console.WriteLine($"  Error processing worksheet \"{sheet.Name}\": {wsEx.Message}");
                    }
                }

                // Save the workbook with charts
                Console.WriteLine($"Saving workbook with charts to \"{outputPath}\"...");
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine("Process completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during processing:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }

        private static Workbook CreateSampleWorkbook()
        {
            Workbook wb = new Workbook();

            // Worksheet 1
            Worksheet ws1 = wb.Worksheets[0];
            ws1.Name = "Fruits";
            ws1.Cells["A1"].PutValue("Category");
            ws1.Cells["B1"].PutValue("Quantity");
            ws1.Cells["A2"].PutValue("Apple");
            ws1.Cells["B2"].PutValue(30);
            ws1.Cells["A3"].PutValue("Banana");
            ws1.Cells["B3"].PutValue(45);
            ws1.Cells["A4"].PutValue("Cherry");
            ws1.Cells["B4"].PutValue(25);

            // Worksheet 2
            Worksheet ws2 = wb.Worksheets.Add("Vehicles");
            ws2.Cells["A1"].PutValue("Category");
            ws2.Cells["B1"].PutValue("Units");
            ws2.Cells["A2"].PutValue("Car");
            ws2.Cells["B2"].PutValue(15);
            ws2.Cells["A3"].PutValue("Bike");
            ws2.Cells["B3"].PutValue(40);
            ws2.Cells["A4"].PutValue("Truck");
            ws2.Cells["B4"].PutValue(10);

            return wb;
        }
    }
}
