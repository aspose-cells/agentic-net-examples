using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

public static class WorkbookProcessor
{
    /// <summary>
    /// Loads a workbook from the given byte array, modifies the first chart's axes,
    /// and returns the updated workbook as a byte array in XLSX format.
    /// </summary>
    public static byte[] ModifyChartAxes(byte[] inputBytes)
    {
        try
        {
            using (MemoryStream inputStream = new MemoryStream(inputBytes))
            {
                Workbook workbook = new Workbook(inputStream);
                Worksheet worksheet = workbook.Worksheets[0];

                if (worksheet.Charts.Count > 0)
                {
                    Chart chart = worksheet.Charts[0];
                    chart.CategoryAxis.Title.Text = "Modified Category Axis";
                    chart.ValueAxis.Title.Text = "Modified Value Axis";
                    chart.ValueAxis.MinValue = 0;
                    chart.ValueAxis.MaxValue = 100;
                }

                using (MemoryStream outputStream = new MemoryStream())
                {
                    workbook.Save(outputStream, SaveFormat.Xlsx);
                    return outputStream.ToArray();
                }
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error modifying chart axes: {ex.Message}");
            return null;
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        // Expect input and output file paths as arguments
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: AsposeCellsRunner <inputFilePath> <outputFilePath>");
            return;
        }

        string inputPath = args[0];
        string outputPath = args[1];

        try
        {
            if (!File.Exists(inputPath))
            {
                Console.Error.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            byte[] inputBytes = File.ReadAllBytes(inputPath);
            byte[] resultBytes = WorkbookProcessor.ModifyChartAxes(inputBytes);

            if (resultBytes == null)
            {
                Console.Error.WriteLine("Processing failed.");
                return;
            }

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            File.WriteAllBytes(outputPath, resultBytes);
            Console.WriteLine($"Modified workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}