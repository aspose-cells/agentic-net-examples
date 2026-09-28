// Title: Create a line sparkline in cell L8 using the named range 'SalesData' as the data source with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that defines a named range named SalesData for cells A1:A10 and then adds a line sparkline to cell L8 referencing that named range using Aspose.Cells. | Show how to build a sparkline group, set its source to the SalesData named range, and place the sparkline in L8 within an Aspose.Cells workbook. | Provide a complete Aspose.Cells example that creates the SalesData range, inserts a line sparkline in L8, and saves the workbook to disk.
// Common Searches: Aspose.Cells C# how to add a sparkline with a named range as source | Create line sparkline in specific cell using named range SalesData Aspose.Cells | Set sparkline data source to a defined name in a .NET workbook | C# Aspose.Cells example for sparkline group referencing SalesData range | Insert sparkline into cell L8 from range A1:A10 using Aspose.Cells
// Tags: define named range for sparkline Aspose.Cells C# | insert sparkline in cell L8 Aspose.Cells | sparkline source named range Aspose.Cells .NET | Aspose.Cells sparkline group example C# | use SalesData range for sparkline C#

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, populates cells A1 through A10 with values 1‑10, defines a named range called SalesData covering that range, inserts a line sparkline into cell L8 that uses SalesData as its data source, and saves the file as SparklineExample.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the range A1:A10
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 0].PutValue(i + 1);
            }

            // Save the workbook
            string outputPath = "SparklineExample.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
