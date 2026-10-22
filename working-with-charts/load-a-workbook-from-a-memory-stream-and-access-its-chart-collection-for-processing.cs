// Title: Load an Excel workbook from a MemoryStream and enumerate all charts in each worksheet using Aspose.Cells for .NET
// AI Prompts: Read an Excel file into a byte array, create a MemoryStream, and use Aspose.Cells to list every chart's name, type, and position across all worksheets. | Open a workbook via a MemoryStream with Aspose.Cells, then loop through each sheet's ChartCollection to output the chart index, title, and chart kind. | When the target .xlsx file is absent, generate a simple workbook, load it from memory, and extract chart metadata using Aspose.Cells in C#.
// Common Searches: aspocells load workbook from memory stream and retrieve chart collection | c# enumerate charts on each worksheet after opening Excel file from byte array | how to list chart names and types using Aspose.Cells with a MemoryStream
// Tags: initialize workbook from MemoryStream Aspose.Cells | enumerate chart collection per worksheet .NET | extract chart metadata (name, type) C# | generate minimal workbook when missing Aspose.Cells | read Excel file bytes into MemoryStream

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example ensures a sample Excel file exists, reads it into a byte array, loads the workbook from a MemoryStream, then iterates each worksheet's ChartCollection, printing worksheet name, chart index, chart name, and chart type.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string filePath = "SampleWorkbook.xlsx";

                // Ensure the Excel file exists; create a minimal workbook if missing.
                if (!File.Exists(filePath))
                {
                    var wb = new Workbook();
                    wb.Worksheets[0].Name = "Sheet1";
                    wb.Save(filePath);
                }

                // Load the workbook bytes safely.
                byte[] excelData = File.ReadAllBytes(filePath);

                // Load the workbook from a memory stream.
                using (MemoryStream memoryStream = new MemoryStream(excelData))
                {
                    Workbook workbook = new Workbook(memoryStream);

                    // Iterate through each worksheet.
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Access the chart collection of the current worksheet.
                        ChartCollection charts = sheet.Charts;

                        // Process each chart in the collection using index-based loop.
                        for (int i = 0; i < charts.Count; i++)
                        {
                            Chart chart = charts[i];
                            Console.WriteLine($"Worksheet: {sheet.Name}");
                            Console.WriteLine($"Chart Index: {i}");
                            Console.WriteLine($"Chart Name: {chart.Name}");
                            Console.WriteLine($"Chart Type: {chart.Type}");
                            Console.WriteLine(new string('-', 40));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
