// Title: Load an Excel workbook from a byte array, modify chart axis titles and scaling, and return the updated workbook as a byte array using Aspose.Cells for .NET
// AI Prompts: Write C# code that reads an XLSX file from a byte[] with Aspose.Cells, iterates every worksheet and chart, sets a custom title, minimum, maximum and major unit for the value axis, updates the category axis title, and returns the modified workbook as a byte[]. | Create a C# method that accepts a byte array containing an Excel workbook, applies axis property changes (title, range, tick interval) to all charts, saves the workbook to a MemoryStream in XLSX format, and outputs the resulting byte array.
// Common Searches: asp.net read excel from byte array and edit chart axes using Aspose.Cells | c# change chart value axis min and max in an in‑memory workbook with Aspose.Cells | how to set custom titles for chart axes when working with a workbook loaded from a stream | save modified Excel workbook to byte[] after updating chart properties in C#
// Tags: Aspose.Cells modify chart axes | C# load workbook from byte array | Aspose.Cells set chart axis title | Aspose.Cells set axis range | Aspose.Cells save workbook to byte array | Aspose.Cells iterate worksheets charts

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an Excel file from a byte array, walks through each worksheet and its charts, customizes the value and category axis titles and scaling, then saves the workbook back to a MemoryStream and returns the updated content as a byte array in XLSX format.
public class ChartAxisModifier
{
    // Loads a workbook from a byte array, modifies chart axes, and returns the updated workbook as a byte array.
    public static byte[] ModifyChartAxes(byte[] inputBytes)
    {
        try
        {
            using (MemoryStream inputStream = new MemoryStream(inputBytes))
            {
                Workbook workbook = new Workbook(inputStream);

                // Iterate through all worksheets in the workbook.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all charts on the current worksheet.
                    foreach (Chart chart in sheet.Charts)
                    {
                        // ----- Modify the Value (Y) Axis -----
                        Axis valueAxis = chart.ValueAxis;
                        if (valueAxis != null && valueAxis.Title != null)
                        {
                            // Set a custom title for the Y axis.
                            valueAxis.Title.Text = "Modified Y Axis";

                            // Define the visible range of the axis.
                            valueAxis.MinValue = 0;
                            valueAxis.MaxValue = 100;

                            // Set the interval between major tick marks.
                            valueAxis.MajorUnit = 10;
                        }

                        // ----- Modify the Category (X) Axis -----
                        Axis categoryAxis = chart.CategoryAxis;
                        if (categoryAxis != null && categoryAxis.Title != null)
                        {
                            // Set a custom title for the X axis.
                            categoryAxis.Title.Text = "Modified X Axis";

                            // Rotate the category labels for better readability.
                            // Note: TickLabelRotationAngle is not available in this version of Aspose.Cells.
                            // If needed, upgrade the library or use alternative formatting options.
                        }
                    }
                }

                // Save the modified workbook to a memory stream and return its byte array.
                using (MemoryStream outputStream = new MemoryStream())
                {
                    workbook.Save(outputStream, SaveFormat.Xlsx);
                    return outputStream.ToArray();
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error modifying chart axes: " + ex.Message);
            return null;
        }
    }

    // Entry point for demonstration.
    public static void Main(string[] args)
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            byte[] inputBytes;

            if (File.Exists(inputPath))
            {
                // Load existing workbook.
                inputBytes = File.ReadAllBytes(inputPath);
            }
            else
            {
                // Create a sample workbook with a chart if the input file does not exist.
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];
                ws.Cells["A1"].PutValue("Category");
                ws.Cells["B1"].PutValue("Value");
                ws.Cells["A2"].PutValue("A");
                ws.Cells["B2"].PutValue(10);
                ws.Cells["A3"].PutValue("B");
                ws.Cells["B3"].PutValue(30);
                ws.Cells["A4"].PutValue("C");
                ws.Cells["B4"].PutValue(70);

                // Add a column chart.
                int chartIndex = ws.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = ws.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                using (MemoryStream ms = new MemoryStream())
                {
                    wb.Save(ms, SaveFormat.Xlsx);
                    inputBytes = ms.ToArray();
                }
            }

            // Modify chart axes.
            byte[] resultBytes = ModifyChartAxes(inputBytes);
            if (resultBytes != null)
            {
                File.WriteAllBytes(outputPath, resultBytes);
                Console.WriteLine($"Modified workbook saved to '{outputPath}'.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Unexpected error: " + ex.Message);
        }
    }
}
