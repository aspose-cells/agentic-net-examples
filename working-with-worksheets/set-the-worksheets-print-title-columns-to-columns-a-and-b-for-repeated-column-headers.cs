// Title: How to repeat columns A and B as print titles in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set PrintTitleColumns to A:B | repeat first two columns as header rows when printing with Aspose.Cells | PageSetup.PrintTitleColumns example in C# Aspose.Cells | configure worksheet print titles for specific columns using Aspose.Cells .NET | set print titles for columns A and B in an Excel workbook with Aspose.Cells
// Tags: Aspose.Cells PageSetup.PrintTitleColumns | C# worksheet print title columns | Excel worksheet page setup Aspose.Cells | set print titles for columns A B .NET | Aspose.Cells repeat columns as print titles

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // The example creates a new Workbook, accesses the first Worksheet, assigns "A:B" to the PageSetup.PrintTitleColumns property to repeat columns A and B as print titles, ensures the output directory exists, saves the workbook as Output.xlsx, and logs a success message.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet (index 0)
                Worksheet worksheet = workbook.Worksheets[0];

                // Set the print title columns to columns A and B (repeated column headers)
                // PageSetup.PrintTitleColumns expects a string range, e.g., "A:B"
                worksheet.PageSetup.PrintTitleColumns = "A:B";

                // Define output file path
                string outputPath = "Output.xlsx";

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
