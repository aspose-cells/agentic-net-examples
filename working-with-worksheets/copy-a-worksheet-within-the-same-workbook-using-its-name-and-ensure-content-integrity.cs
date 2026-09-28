// Title: Copy a worksheet by its name within the same workbook and verify data integrity using Aspose.Cells for .NET
// AI Prompts: Locate a worksheet by its name with Aspose.Cells, duplicate it in the same workbook, rename the copy, and save the workbook. | Loop through the used range of the original and copied worksheets to compare each cell value and ensure the copy matches the source. | Add robust error handling for missing input files or nonexistent worksheet names when performing a worksheet copy with Aspose.Cells.
// Common Searches: Aspose.Cells C# copy worksheet within workbook by sheet name | how to verify that a duplicated sheet has identical data using Aspose.Cells | C# example of Aspose.Cells AddCopy method with data integrity check | duplicate Excel sheet and compare cell values Aspose.Cells .NET
// Tags: Aspose.Cells AddCopy worksheet duplication | copy worksheet by name .NET | worksheet data integrity verification Aspose.Cells | compare cell values after sheet copy C# | save workbook after sheet duplication Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace WorksheetCopyExample
{
    // The example loads an existing Excel file, retrieves a worksheet by its name, creates a copy of that worksheet within the same workbook, renames the copy, iterates over all used cells to confirm that the copied data matches the original, and finally saves the workbook, handling missing files and missing worksheets gracefully.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    throw new FileNotFoundException($"Input file not found: {inputPath}");
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Retrieve the source worksheet by name
                string sourceSheetName = "SourceSheetName";
                Worksheet sourceSheet = workbook.Worksheets[sourceSheetName];

                if (sourceSheet == null)
                {
                    throw new Exception($"Worksheet '{sourceSheetName}' not found in the workbook.");
                }

                // Copy the worksheet within the same workbook
                int copiedIndex = workbook.Worksheets.AddCopy(sourceSheet.Index);
                Worksheet copiedSheet = workbook.Worksheets[copiedIndex];
                copiedSheet.Name = sourceSheetName + "_Copy";

                // Verify that the copied content matches the source
                bool contentIsIdentical = true;
                int maxRow = sourceSheet.Cells.MaxDataRow;
                int maxCol = sourceSheet.Cells.MaxDataColumn;

                for (int row = 0; row <= maxRow && contentIsIdentical; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        object sourceValue = sourceSheet.Cells[row, col].Value;
                        object copiedValue = copiedSheet.Cells[row, col].Value;

                        if (!object.Equals(sourceValue, copiedValue))
                        {
                            contentIsIdentical = false;
                            break;
                        }
                    }
                }

                if (!contentIsIdentical)
                {
                    throw new Exception("Worksheet copy integrity check failed.");
                }

                // Save the workbook to the desired output path
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
