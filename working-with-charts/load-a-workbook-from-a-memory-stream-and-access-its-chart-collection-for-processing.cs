// Title: Load an Excel workbook from a MemoryStream and list all worksheet charts using Aspose.Cells for .NET
// Common Searches: how to read an Excel file into a MemoryStream and access its charts with Aspose.Cells C# | enumerate chart objects on each worksheet after loading workbook from byte array Aspose.Cells | Aspose.Cells example for iterating chart collection from a workbook loaded from memory
// Tags: load workbook from memory stream Aspose.Cells | iterate worksheet chart collection C# | access charts in Excel workbook Aspose.Cells | process chart objects from byte array Aspose.Cells | read Excel file into byte array C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample reads an Excel file into a byte array, creates a MemoryStream, loads it into an Aspose.Cells Workbook, then iterates through each worksheet's ChartCollection, outputting the worksheet name, chart index, type, and name while handling potential exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the Excel file into a byte array.
            byte[] excelData = File.ReadAllBytes(inputPath);

            // Load a workbook from a memory stream.
            using (MemoryStream ms = new MemoryStream(excelData))
            {
                Workbook workbook = new Workbook(ms);

                // Iterate through all worksheets in the workbook.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Access the chart collection of the current worksheet.
                    ChartCollection charts = sheet.Charts;

                    // Process each chart in the collection.
                    for (int i = 0; i < charts.Count; i++)
                    {
                        Chart chart = charts[i];
                        Console.WriteLine($"Worksheet: {sheet.Name}, Chart Index: {i}, Type: {chart.Type}, Name: {chart.Name}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message.
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
