// Title: How to hide column B in an Excel workbook by selecting its EntireColumn range using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide specific column using EntireColumn range | programmatically hide column B in an Excel file with Aspose.Cells .NET | C# example to set column visibility to hidden in a generated workbook | using HideColumn method to conceal a column in Aspose.Cells | control Excel column visibility from code with Aspose.Cells for .NET
// Tags: Aspose.Cells hide column C# | EntireColumn range visibility Aspose.Cells | Excel column B hidden Aspose.Cells .NET | Workbook column visibility control Aspose.Cells | C# hide Excel column programmatically

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new Workbook, accesses the first Worksheet, selects column B via its index, hides it using the HideColumn method, ensures the output directory exists, saves the file as output.xlsx, and includes basic exception handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Hide column B (zero‑based index 1)
            // Use HideColumn method which is available in all Aspose.Cells versions
            sheet.Cells.HideColumn(1);

            // Define output file path
            string outputPath = "output.xlsx";

            // Ensure the directory for the output file exists
            string fullPath = Path.GetFullPath(outputPath);
            string? outputDir = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
