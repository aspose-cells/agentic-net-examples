// Title: Process Excel smart markers from a byte array and return the updated workbook as a byte array using Aspose.Cells in C#
// AI Prompts: Generate a C# method that accepts a byte[] with an XLSX file, invokes WorkbookDesigner to evaluate all smart markers, and returns the modified workbook as a new byte[]. | Create a utility class that loads an Excel workbook from a MemoryStream, runs smart‑marker processing, and writes the result to another MemoryStream without using the file system.
// Common Searches: Aspose.Cells smart marker evaluation from byte array C# | C# convert processed Excel workbook to byte[] without saving to disk | How to use WorkbookDesigner with streams to obtain a byte array result | Processing smart markers in memory using Aspose.Cells API | Example of using WorkbookDesigner without file I/O
// Tags: WorkbookDesigner memory stream processing | Aspose.Cells load workbook from byte array | save workbook to byte array XLSX | C# in‑memory smart marker handling | avoid temporary files Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// // Loads an Excel workbook from a byte array, processes all smart markers with WorkbookDesigner, and returns the updated workbook as a byte array.
public static class SmartMarkerHelper
{
    /// <param name="inputData">The input workbook data (e.g., an XLSX file) as a byte array.</param>
    /// <returns>The processed workbook data as a byte array.</returns>
    public static byte[] ProcessSmartMarkers(byte[] inputData)
    {
        try
        {
            // Load the workbook from the input memory stream
            using (var inputStream = new MemoryStream(inputData))
            {
                var workbook = new Workbook(inputStream);

                // Process all smart markers present in the workbook using WorkbookDesigner
                var designer = new WorkbookDesigner(workbook);
                designer.Process();

                // Save the processed workbook to an output memory stream
                using (var outputStream = new MemoryStream())
                {
                    workbook.Save(outputStream, SaveFormat.Xlsx);
                    return outputStream.ToArray();
                }
            }
        }
        catch (Exception ex)
        {
            // Log or handle the exception as needed
            Console.Error.WriteLine($"Error processing smart markers: {ex.Message}");
            throw;
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            byte[] inputData = File.ReadAllBytes(inputPath);
            byte[] resultData = SmartMarkerHelper.ProcessSmartMarkers(inputData);
            File.WriteAllBytes(outputPath, resultData);
            Console.WriteLine($"Processed workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
