using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsChartFilteringDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ------------------------------------------------------------
                // Sample data for two charts
                // ------------------------------------------------------------
                // Header
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["D1"].PutValue("Series3");

                // Data rows
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["C2"].PutValue(0);   // All zeros for Series2
                sheet.Cells["D2"].PutValue(5);

                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["C3"].PutValue(0);   // All zeros for Series2
                sheet.Cells["D3"].PutValue(15);

                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["C4"].PutValue(0);   // All zeros for Series2
                sheet.Cells["D4"].PutValue(25);

                // ------------------------------------------------------------
                // Add first chart (Column)
                // ------------------------------------------------------------
                int chartIdx1 = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
                Chart chart1 = sheet.Charts[chartIdx1];
                chart1.NSeries.Add("B2:B4", true); // Series1
                chart1.NSeries.Add("C2:C4", true); // Series2 (all zeros)
                chart1.NSeries.Add("D2:D4", true); // Series3
                chart1.NSeries.CategoryData = "A2:A4";

                // ------------------------------------------------------------
                // Add second chart (Line)
                // ------------------------------------------------------------
                int chartIdx2 = sheet.Charts.Add(ChartType.Line, 6, 13, 20, 25);
                Chart chart2 = sheet.Charts[chartIdx2];
                chart2.NSeries.Add("B2:B4", true); // Series1
                chart2.NSeries.Add("C2:C4", true); // Series2 (all zeros)
                chart2.NSeries.Add("D2:D4", true); // Series3
                chart2.NSeries.CategoryData = "A2:A4";

                // ------------------------------------------------------------
                // Hide series that contain only zero values across all charts
                // ------------------------------------------------------------
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    foreach (Chart ch in ws.Charts)
                    {
                        foreach (Series ser in ch.NSeries)
                        {
                            // Get the range string that defines the series values (e.g., "B2:B4")
                            string valueRange = ser.Values;

                            // Skip if the range is empty
                            if (string.IsNullOrEmpty(valueRange))
                                continue;

                            // Create a Range object from the address string
                            AsposeRange range = ws.Cells.CreateRange(valueRange);
                            bool allZero = true;

                            // Iterate through each cell in the range
                            for (int r = 0; r < range.RowCount && allZero; r++)
                            {
                                for (int c = 0; c < range.ColumnCount && allZero; c++)
                                {
                                    Cell cell = range[r, c];
                                    double cellValue = 0;

                                    if (cell.Type == CellValueType.IsNumeric)
                                        cellValue = cell.DoubleValue;
                                    else if (cell.Type == CellValueType.IsString && double.TryParse(cell.StringValue, out double parsed))
                                        cellValue = parsed;

                                    if (cellValue != 0)
                                        allZero = false;
                                }
                            }

                            // Hide the series if all its values are zero
                            if (allZero)
                                ser.IsFiltered = true;
                        }
                    }
                }

                // ------------------------------------------------------------
                // Save the workbook
                // ------------------------------------------------------------
                string outputPath = "ChartsWithZeroSeriesFiltered.xlsx";

                // Ensure the directory exists (handle possible null from GetDirectoryName)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    Directory.CreateDirectory(outputDir);

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}