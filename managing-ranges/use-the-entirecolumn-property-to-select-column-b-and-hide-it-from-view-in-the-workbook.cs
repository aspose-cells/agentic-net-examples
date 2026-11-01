// Title: Hide column B in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells hide specific column in Excel worksheet | How to programmatically hide column B with Aspose.Cells .NET | Aspose.Cells hide column by index example C# | Set column visibility to hidden using Aspose.Cells in C# | Hide Excel column using Cells.HideColumn method Aspose.Cells
// Tags: Aspose.Cells hide column | C# hide Excel column programmatically | Cells.HideColumn method .NET | Excel column visibility Aspose.Cells | hide column B Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, accesses the first Worksheet, and hides column B (zero‑based index 1) with sheet.Cells.HideColumn(1). It ensures the output directory exists, saves the file as output.xlsx, and handles any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Hide column B (zero‑based index 1)
            // Use Cells collection to access column properties
            sheet.Cells.HideColumn(1);

            // Define output file path
            string outputPath = "output.xlsx";

            // Determine the output directory (handle case when there is no directory part)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? Directory.GetCurrentDirectory();

            // Ensure the output directory exists
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
