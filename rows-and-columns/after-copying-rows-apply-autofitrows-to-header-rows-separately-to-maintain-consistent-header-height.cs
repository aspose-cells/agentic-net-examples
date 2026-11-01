// Title: Copy rows 3‑11 to a new worksheet and auto‑fit only the header rows using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that copies rows 3 through 11 from the first worksheet to a newly created sheet and then calls AutoFitRows on rows 0‑1 only, preserving the heights of all other rows. | Generate a .NET program that loads an Excel workbook, copies a specific block of rows to another worksheet, and uses AutoFitterOptions to auto‑size just the header rows.
// Common Searches: Aspose.Cells copy a range of rows to another sheet and auto fit only the first two rows in C# | How to preserve row heights while copying rows with Aspose.Cells .NET | AutoFitRows for header rows after copying rows using Aspose.Cells | C# Aspose.Cells copy a block of rows to a new worksheet and auto‑fit header rows
// Tags: copy rows between worksheets Aspose.Cells C# | AutoFitRows header rows Aspose.Cells | AutoFitterOptions usage Aspose.Cells | preserve row height Aspose.Cells copy | save workbook after row manipulation Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The program loads "input.xlsx", copies rows 3‑11 from the first worksheet to a newly added sheet named "CopiedRows", applies AutoFitRows only to the first two rows (header) using AutoFitterOptions, and saves the modified workbook as "output.xlsx".
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Source worksheet (first sheet)
            Worksheet sourceSheet = workbook.Worksheets[0];

            // Destination worksheet (add a new sheet)
            Worksheet destSheet = workbook.Worksheets.Add("CopiedRows");

            // Define the range of rows to copy from the source sheet
            int sourceStartRow = 2;          // zero‑based index, e.g., row 3 in Excel
            int rowsToCopy = 9;              // number of rows to copy (rows 3‑11)
            int destStartRow = 0;            // where to paste in the destination sheet

            // Copy rows from source to destination
            destSheet.Cells.CopyRows(sourceSheet.Cells, sourceStartRow, rowsToCopy, destStartRow);

            // AutoFit only the header rows (rows 0 and 1) in the destination sheet
            var autofitOptions = new AutoFitterOptions(); // default options
            // totalRows = 2 because we want rows 0 and 1
            destSheet.AutoFitRows(0, 2, autofitOptions);

            // Save the modified workbook
            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
