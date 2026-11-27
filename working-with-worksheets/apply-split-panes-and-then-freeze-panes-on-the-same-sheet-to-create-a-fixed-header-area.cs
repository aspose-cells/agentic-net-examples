// Title: How to freeze the top five rows as a fixed header in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# freeze first 5 rows in worksheet | Create a fixed header area with FreezePanes method in .NET | Freeze top rows while scrolling using Aspose.Cells FreezePanes overload | Save Excel file after applying freeze panes with Aspose.Cells | Ensure output directory exists before saving workbook Aspose.Cells C#
// Tags: freeze panes top rows Aspose.Cells | fixed header Excel .NET | FreezePanes overload four arguments | save workbook with frozen header Aspose.Cells | output directory creation Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a new Workbook, accesses the first Worksheet, freezes the first five rows using the four‑argument FreezePanes method to create a fixed header, ensures the target directory exists, and saves the file as FixedHeader.xlsx while handling potential exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Freeze the top 5 rows (the header area) so they stay visible while scrolling
                // Overload with four arguments to ensure compatibility across versions
                sheet.FreezePanes(5, 0, 5, 0);

                // Define output file path
                string outputPath = "FixedHeader.xlsx";

                // Ensure the directory exists (if a directory part is present)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
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
                Console.WriteLine("An error occurred while creating the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
