// Title: Freeze rows and columns in an Excel worksheet and write the workbook to a MemoryStream using Aspose.Cells for .NET
// AI Prompts: Freeze the first two rows and three columns of a worksheet, then save the workbook as an Xlsx stream with Aspose.Cells. | Update the example to accept variable row and column counts for FreezePanes and output the resulting workbook to a MemoryStream.
// Common Searches: Aspose.Cells how to freeze specific rows and columns and get a memory stream | C# save Excel workbook to MemoryStream after applying FreezePanes with Aspose.Cells | example of using FreezePanes method and streaming the workbook in .NET
// Tags: Aspose.Cells FreezePanes usage | freeze rows columns Aspose.Cells .NET | save workbook to MemoryStream Aspose.Cells | export frozen worksheet to Xlsx stream | dynamic FreezePanes parameters C#

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates creating a new workbook, freezing the first two rows and three columns on the first worksheet using FreezePanes, and saving the workbook into a MemoryStream in Xlsx format with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            using (var workbook = new Workbook())
            {
                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Freeze the first 2 rows and the first 3 columns
                // row = 2 (first row after frozen rows), column = 3 (first column after frozen columns)
                // totalRows = 2, totalColumns = 3 specify how many rows/columns to freeze
                sheet.FreezePanes(2, 3, 2, 3);

                // Save the workbook to a memory stream
                using (var memoryStream = new MemoryStream())
                {
                    workbook.Save(memoryStream, SaveFormat.Xlsx);
                    memoryStream.Position = 0;

                    // Example: write to a file (optional)
                    // File.WriteAllBytes("output.xlsx", memoryStream.ToArray());
                }
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
