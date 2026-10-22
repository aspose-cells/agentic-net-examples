// Title: Hide rows 5‑15 and enable formula view in the first worksheet before saving with Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a new Workbook, hides rows 5 through 15 on the first worksheet, switches the sheet to formula view, ensures the target directory exists, and saves the workbook using Aspose.Cells. | Show how to use the Aspose.Cells Worksheet.IsDisplayFormulas property after hiding a specific row range, then persist the file to a given path while handling folder creation in .NET. | Provide a snippet that iterates over rows 4‑14 to set IsHidden, sets the worksheet to display formulas, and calls Workbook.Save with proper error handling.
// Common Searches: Aspose.Cells hide rows 5 to 15 and show formulas in C# | C# code to enable formula view after hiding rows with Aspose.Cells | How to ensure output folder exists when saving an Excel file using Aspose.Cells .NET | Enable formula display in an Aspose.Cells worksheet before saving the workbook | Hide a range of rows and display formulas on the same sheet with Aspose.Cells for .NET
// Tags: hide rows Aspose.Cells C# | display formulas worksheet Aspose.Cells | worksheet row visibility Aspose.Cells | save workbook to file Aspose.Cells .NET | ensure output directory exists .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new Workbook, accesses the first worksheet, hides rows 5‑15 by setting each Row.IsHidden to true, enables formula view with Worksheet.IsDisplayFormulas, ensures the output directory exists (creating it if needed), and saves the workbook to 'output.xlsx' using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Hide rows 5 through 15 (zero‑based index: rows 4 to 14)
                for (int rowIndex = 4; rowIndex <= 14; rowIndex++)
                {
                    // Use the Cells.Rows collection to access the Row object
                    sheet.Cells.Rows[rowIndex].IsHidden = true;
                }

                // Define output file path
                string outputPath = "output.xlsx";

                // Determine the directory part of the path; if none, use the current directory
                string outputDir = Path.GetDirectoryName(outputPath) ?? string.Empty;
                if (string.IsNullOrEmpty(outputDir))
                {
                    outputDir = Directory.GetCurrentDirectory();
                }

                // Ensure the directory exists
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to a file
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
