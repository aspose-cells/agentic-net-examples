// Title: Copy a template row to multiple rows in every worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Use Cells.CopyRows to copy row 2 from the first sheet and paste it into rows 5, 10, and 15 of each worksheet in a workbook. | Loop through all worksheets in a workbook and duplicate a template row to several destination rows with a nested loop in C#. | Load a workbook, verify the template file exists, copy the template row to multiple target rows across sheets, and save the result to a new file using Aspose.Cells.
// Common Searches: aspnet copy a single row to several rows in each sheet using Aspose.Cells | c# Aspose.Cells copy template row to multiple rows across worksheets | how to duplicate a row in every worksheet with Aspose.Cells loop | batch copy rows in Excel workbook using Aspose.Cells C# example | copy rows to specific row numbers in all sheets Aspose.Cells
// Tags: batch row copy Cells.CopyRows Aspose.Cells | template row replication across worksheets C# | worksheet loop copy rows Aspose.Cells | save modified workbook Aspose.Cells | handle missing template file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example loads a template.xlsx workbook, copies the row at index 1 from the first worksheet to rows 5, 10, and 15 of every worksheet using Cells.CopyRows inside nested loops, and saves the updated workbook as output.xlsx while handling missing‑file errors.
    public class BatchCopyTemplateRow
    {
        public static void Run()
        {
            const string templatePath = "template.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the template file exists to avoid FileNotFoundException.
                if (!File.Exists(templatePath))
                {
                    throw new FileNotFoundException($"The template file '{templatePath}' was not found.");
                }

                // Load the workbook that contains the template row.
                Workbook workbook = new Workbook(templatePath);

                // Define the source worksheet and the row index of the template row (0‑based).
                Worksheet sourceSheet = workbook.Worksheets[0];
                int templateRowIndex = 1; // e.g., row 2 in Excel

                // Define target rows for each worksheet.
                // For demonstration, each worksheet will receive the template row at rows 5, 10 and 15.
                int[] targetRowIndices = new int[] { 4, 9, 14 }; // 0‑based indices

                // Loop through all worksheets in the workbook.
                foreach (Worksheet targetSheet in workbook.Worksheets)
                {
                    Cells sourceCells = sourceSheet.Cells;
                    Cells targetCells = targetSheet.Cells;

                    // Copy the template row to each target position within the current worksheet.
                    foreach (int destRowIndex in targetRowIndices)
                    {
                        // Copy a single row (rowNumber = 1) from the source worksheet to the destination worksheet.
                        targetCells.CopyRows(sourceCells, templateRowIndex, destRowIndex, 1);
                    }
                }

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log the exception details.
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }

    // Entry point for the application.
    public class Program
    {
        public static void Main(string[] args)
        {
            BatchCopyTemplateRow.Run();
        }
    }
}
