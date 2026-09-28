// Title: Load an Excel workbook from a byte array, modify chart axis titles and value range, and save back to a byte array using Aspose.Cells for .NET
// AI Prompts: Read an Excel file from a byte[] via MemoryStream, loop through every worksheet and chart, set custom text for CategoryAxis.Title and ValueAxis.Title, define ValueAxis.MinValue and ValueAxis.MaxValue, then write the workbook to a new MemoryStream and return the resulting byte array with Aspose.Cells in C#. | Load a workbook from a byte array, adjust each chart's axis labels and numeric limits, and export the updated workbook as a byte[] without touching the file system, using the Aspose.Cells chart API.
// Common Searches: Aspose.Cells C# change chart axis title from byte array | How to set chart value axis min and max programmatically with Aspose.Cells | Save modified Excel workbook to byte[] after editing charts in .NET | Iterate over all charts in a workbook loaded from MemoryStream using Aspose.Cells | Update chart axis labels in an in-memory Excel file with Aspose.Cells
// Tags: Aspose.Cells chart axis title update | byte array workbook loading Aspose.Cells | chart value axis min max Aspose.Cells | enumerate charts worksheet Aspose.Cells | export workbook to byte array .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads an Excel workbook from a byte array using a MemoryStream, iterates through each worksheet and its charts, changes the CategoryAxis and ValueAxis titles, sets a fixed minimum and maximum for the value axis, and then saves the modified workbook to another MemoryStream, returning the updated workbook as a byte array.
public class ChartAxisModifier
{
    // Loads a workbook from a byte array, modifies chart axes, and returns the updated workbook as a byte array.
    public static byte[] ModifyChartAxes(byte[] sourceBytes)
    {
        try
        {
            using (MemoryStream inputStream = new MemoryStream(sourceBytes))
            using (Workbook workbook = new Workbook(inputStream))
            {
                // Iterate through all worksheets
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    // Iterate through all charts in the worksheet
                    foreach (Chart chart in worksheet.Charts)
                    {
                        // Modify the X (category) axis title
                        chart.CategoryAxis.Title.Text = "Modified Category Axis";

                        // Modify the Y (value) axis title
                        chart.ValueAxis.Title.Text = "Modified Value Axis";

                        // Example: set a fixed range for the value axis
                        chart.ValueAxis.MinValue = 0;
                        chart.ValueAxis.MaxValue = 100;
                    }
                }

                // Save the modified workbook to an output stream
                using (MemoryStream outputStream = new MemoryStream())
                {
                    workbook.Save(outputStream, SaveFormat.Xlsx);
                    return outputStream.ToArray();
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error modifying chart axes: {ex.Message}");
            throw;
        }
    }

    // Demonstration of usage
    public static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists; if not, create a simple workbook with a chart
            if (!File.Exists(inputPath))
            {
                CreateSampleWorkbookWithChart(inputPath);
                Console.WriteLine($"Sample workbook created at '{inputPath}'.");
            }

            // Load an existing Excel file into a byte array
            byte[] originalBytes = File.ReadAllBytes(inputPath);

            // Modify chart axes and obtain the updated workbook bytes
            byte[] modifiedBytes = ModifyChartAxes(originalBytes);

            // Write the modified workbook to a file for verification
            File.WriteAllBytes(outputPath, modifiedBytes);

            Console.WriteLine($"Chart axes have been modified and saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }

    // Helper method to create a sample workbook containing a simple chart
    private static void CreateSampleWorkbookWithChart(string filePath)
    {
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        ws.Name = "Data";

        // Populate sample data
        ws.Cells["A1"].PutValue("Category");
        ws.Cells["B1"].PutValue("Value");
        ws.Cells["A2"].PutValue("A");
        ws.Cells["B2"].PutValue(10);
        ws.Cells["A3"].PutValue("B");
        ws.Cells["B3"].PutValue(40);
        ws.Cells["A4"].PutValue("C");
        ws.Cells["B4"].PutValue(70);

        // Add a chart
        int chartIndex = ws.Charts.Add(ChartType.Column, 5, 0, 15, 7);
        Chart chart = ws.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Save the workbook
        wb.Save(filePath, SaveFormat.Xlsx);
    }
}
