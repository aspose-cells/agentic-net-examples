// Title: Copy a cell range from an existing Excel workbook to a new workbook and set the Author property using Aspose.Cells for .NET
// AI Prompts: Copy the range A1:C10 from Source.xlsx to a new workbook while preserving values and formatting, then save as CopiedRange.xlsx with Aspose.Cells in C#. | Assign the built‑in Author document property of a newly created workbook to "John Doe" using Aspose.Cells for .NET. | Write a manual loop that transfers both cell values and styles from one worksheet to another workbook via Aspose.Cells APIs.
// Common Searches: Aspose.Cells copy specific range to new workbook preserving formatting C# | How to set workbook author metadata with Aspose.Cells .NET | Copy cells A1:C10 from one Excel file to another using Aspose.Cells | Manual range copy example Aspose.Cells C# | Set built‑in document properties in Aspose.Cells workbook
// Tags: Aspose.Cells transfer cell block to new workbook | Aspose.Cells set author built‑in document property | manual cell style replication Aspose.Cells C# | C# preserve formatting during Excel range copy | Aspose.Cells manipulate workbook metadata

using System;
using System.IO;
using Aspose.Cells;

// The program loads Source.xlsx, manually copies the A1:C10 range (including values and styles) to a new workbook, sets the Author built‑in document property to "John Doe", and saves the result as CopiedRange.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source workbook
            const string sourcePath = "Source.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook containing the range to copy
            Workbook sourceWorkbook = new Workbook(sourcePath);
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Define the range to copy (e.g., A1:C10)
            CellArea sourceArea = new CellArea
            {
                StartRow = 0,    // Row 1 (zero‑based)
                StartColumn = 0, // Column A
                EndRow = 9,      // Row 10
                EndColumn = 2    // Column C
            };

            // Create a new workbook (empty) and get its first worksheet
            Workbook newWorkbook = new Workbook();
            Worksheet destSheet = newWorkbook.Worksheets[0];

            // Manually copy the defined range from source to destination
            try
            {
                for (int r = sourceArea.StartRow; r <= sourceArea.EndRow; r++)
                {
                    for (int c = sourceArea.StartColumn; c <= sourceArea.EndColumn; c++)
                    {
                        // Destination cell coordinates relative to the top‑left of the range
                        int destRow = r - sourceArea.StartRow;
                        int destCol = c - sourceArea.StartColumn;

                        // Copy the value and style
                        Cell srcCell = sourceSheet.Cells[r, c];
                        Cell dstCell = destSheet.Cells[destRow, destCol];
                        dstCell.PutValue(srcCell.Value);
                        dstCell.SetStyle(srcCell.GetStyle());
                    }
                }
            }
            catch (Exception copyEx)
            {
                Console.WriteLine($"Error during range copy: {copyEx.Message}");
                return;
            }

            // Set the author property using the strongly‑typed property
            newWorkbook.BuiltInDocumentProperties.Author = "John Doe";

            // Save the new workbook to a file
            const string outputPath = "CopiedRange.xlsx";
            newWorkbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
