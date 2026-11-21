// Title: Hide rows 20‑30 in an Excel worksheet and save the workbook using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells hide rows 20 to 30 in a worksheet | How to conceal a range of rows and save the Excel file with Aspose.Cells .NET | Aspose.Cells set IsHidden property for multiple rows and export workbook | Saving an Excel workbook after hiding specific rows using Aspose.Cells C# | Programmatically hide rows in Excel with Aspose.Cells and write to file
// Tags: hide rows Aspose.Cells | row visibility control Aspose.Cells | save workbook Aspose.Cells | Excel row concealment .NET | Aspose.Cells hide row range

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // The example creates a new Workbook, accesses the first Worksheet, loops through rows 20‑30 (zero‑based indices 19‑29) and sets each row's IsHidden property to true, ensures the output directory exists, then saves the file as HiddenRows.xlsx while handling any exceptions.
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

                // Hide rows 20 to 30 (zero‑based index: 19 to 29)
                for (int rowIndex = 19; rowIndex <= 29; rowIndex++)
                {
                    // Use Cells.Rows collection to hide the row
                    sheet.Cells.Rows[rowIndex].IsHidden = true;
                }

                string outputPath = "HiddenRows.xlsx";

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
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
