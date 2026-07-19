// Title: Modify Excel chart axes from a byte array and return the updated workbook as a byte array using Aspose.Cells for .NET
// Description: Loads an Excel workbook from a byte array via MemoryStream, accesses the first worksheet, checks for a chart, updates the category and value axis titles, tick marks, and sets the value axis range (0‑100). The workbook is then saved to a new MemoryStream and returned as a byte array, with a fallback to the original data on error.
// Keywords: Aspose.Cells | C# | Excel chart axis modification | load workbook from stream | save workbook to byte array | chart category axis title | chart value axis min max | TickMarkType | MemoryStream | .NET Excel automation
// Common Searches: Aspose.Cells change chart axis title programmatically | Load Excel workbook from byte array and edit chart | Set chart value axis minimum and maximum using Aspose.Cells | Modify chart tick marks in a workbook loaded from memory stream | Save modified Excel file to byte array with Aspose.Cells
// Developer Intent: Update the axis titles, tick marks, and value range of the first chart in an Excel workbook loaded from a byte array, then return the modified workbook as a byte array.
// Use Cases: Standardize chart axes in Excel files received via a web API before returning them to the client. | Process uploaded spreadsheets in a serverless function, adjust chart scales for consistency, and store the result as a byte array in a database. | Generate automated reports where chart axes must be fixed (e.g., 0‑100) before attaching the workbook to an email.
// AI Prompts: Create a C# method that loads an Excel workbook from a byte array, changes the first chart's category and value axis titles, sets the value axis range to 0‑100, and returns the updated workbook as a byte array using Aspose.Cells. | Add robust error handling to the chart‑axis‑modification method so that it returns the original byte array on failure and logs the exception details. | Write unit tests for the ModifyChartAxes function that verify the axis titles, tick marks, and min/max values are correctly applied.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an Excel workbook from a byte array via MemoryStream, accesses the first worksheet, checks for a chart, updates the category and value axis titles, tick marks, and sets the value axis range (0‑100). The workbook is then saved to a new MemoryStream and returned as a byte array, with a fallback to the original data on error.
public class ChartAxisModifier
{
    /// <param name="excelData">Input Excel file data.</param>
    /// <returns>Modified Excel file data.</returns>
    public static byte[] ModifyChartAxes(byte[] excelData)
    {
        try
        {
            // Load workbook from byte array
            using (MemoryStream inputStream = new MemoryStream(excelData))
            {
                Workbook workbook = new Workbook(inputStream);
                Worksheet worksheet = workbook.Worksheets[0];

                // Modify first chart if it exists
                if (worksheet.Charts.Count > 0)
                {
                    Chart chart = worksheet.Charts[0];
                    chart.CategoryAxis.Title.Text = "New Category Axis";
                    chart.CategoryAxis.MajorTickMark = TickMarkType.Outside;
                    chart.ValueAxis.Title.Text = "New Value Axis";
                    chart.ValueAxis.MinValue = 0;
                    chart.ValueAxis.MaxValue = 100;
                    chart.ValueAxis.MajorTickMark = TickMarkType.Outside;
                }

                // Save modified workbook to byte array
                using (MemoryStream outputStream = new MemoryStream())
                {
                    workbook.Save(outputStream, SaveFormat.Xlsx);
                    return outputStream.ToArray();
                }
            }
        }
        catch
        {
            // Return original data on failure
            return excelData;
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            byte[] inputData = File.ReadAllBytes(inputPath);
            byte[] resultData = ChartAxisModifier.ModifyChartAxes(inputData);
            File.WriteAllBytes(outputPath, resultData);
            Console.WriteLine($"Modified workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
