// Title: Load an Excel workbook from a MemoryStream, edit chart axes, and obtain the result as a byte array using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# load workbook from MemoryStream and change chart axis titles | How to update X and Y axis settings of a chart after reading an Excel file from a stream in .NET | Save a modified Aspose.Cells workbook to a byte array for further processing | C# example for editing chart axes in a workbook loaded from a stream with Aspose.Cells | Export Excel chart changes to a byte[] using Aspose.Cells without writing to disk
// Tags: memory stream workbook loading Aspose.Cells | chart axis title editing Aspose.Cells | chart axis major/minor unit configuration Aspose.Cells | xlsx workbook export to byte array | chart axis range setting Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartAxisDemo
{
    // The sample creates a workbook with sample data and a column chart, saves it to a MemoryStream, reloads the workbook from that stream, updates the chart's category (X) and value (Y) axes titles, major/minor units, and value range, then saves the modified workbook to another MemoryStream and extracts the XLSX content as a byte array.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // ------------------------------------------------------------
                // 1. Create a sample workbook with data and a chart.
                // ------------------------------------------------------------
                Workbook originalWorkbook = new Workbook();
                Worksheet sheet = originalWorkbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");

                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(180);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 15);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";
                chart.Title.Text = "Quarterly Sales";

                // ------------------------------------------------------------
                // 2. Save the workbook to a memory stream (simulating an input stream).
                // ------------------------------------------------------------
                using (MemoryStream inputStream = new MemoryStream())
                {
                    originalWorkbook.Save(inputStream, SaveFormat.Xlsx);
                    inputStream.Position = 0; // Reset for reading

                    // ------------------------------------------------------------
                    // 3. Load the workbook from the stream.
                    // ------------------------------------------------------------
                    Workbook loadedWorkbook = new Workbook(inputStream);

                    // ------------------------------------------------------------
                    // 4. Modify the chart axes.
                    // ------------------------------------------------------------
                    Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
                    if (loadedSheet.Charts.Count == 0)
                    {
                        Console.WriteLine("No charts found in the loaded workbook.");
                        return;
                    }

                    Chart loadedChart = loadedSheet.Charts[0];

                    try
                    {
                        // Modify X axis (Category axis)
                        Axis xAxis = loadedChart.CategoryAxis;
                        xAxis.Title.Text = "Months";
                        xAxis.Title.IsVisible = true;
                        xAxis.MajorUnit = 1;      // one month per major tick
                        xAxis.MinorUnit = 0.5;

                        // Modify Y axis (Value axis)
                        Axis yAxis = loadedChart.ValueAxis;
                        yAxis.Title.Text = "Sales (Units)";
                        yAxis.Title.IsVisible = true;
                        yAxis.MinorUnit = 10;
                        yAxis.MajorUnit = 50;
                        yAxis.MaxValue = 200;
                        yAxis.MinValue = 0;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error modifying axes: {ex.Message}");
                    }

                    // ------------------------------------------------------------
                    // 5. Save the modified workbook to a new memory stream and obtain a byte array.
                    // ------------------------------------------------------------
                    using (MemoryStream outputStream = new MemoryStream())
                    {
                        loadedWorkbook.Save(outputStream, SaveFormat.Xlsx);
                        byte[] workbookBytes = outputStream.ToArray();

                        Console.WriteLine($"Modified workbook saved to byte array. Size: {workbookBytes.Length} bytes.");

                        // Optional: write the byte array to a file to verify the result.
                        // File.WriteAllBytes("ModifiedWorkbook.xlsx", workbookBytes);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
