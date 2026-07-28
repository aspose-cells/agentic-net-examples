// Title: C# Aspose.Cells example: Unmerge B2:C2, copy its value to C2, and save the workbook
// Description: Loads an Excel file (or creates a sample), reads the value from the merged range B2:C2, unmerges the cells, writes the original value into C2, and saves the result as a new workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells unmerge cells C# | copy merged cell value Aspose.Cells | C# unmerge B2:C2 | Aspose.Cells save workbook after unmerge | .NET Excel unmerge range | Aspose.Cells sample code GitHub | Excel merged cells handling
// Common Searches: how to unmerge a merged range with Aspose.Cells .NET | copy value from merged Excel cells using C# | Aspose.Cells example to split B2:C2 and keep the text | C# code to unmerge cells and duplicate value in adjacent cell | Aspose.Cells unmerge and preserve data
// Developer Intent: Unmerge the merged range B2:C2, duplicate its original value into C2, and save the updated workbook.
// Use Cases: Cleaning user‑uploaded spreadsheets that contain merged headers before importing into a database. | Preparing Excel files for CSV export where merged cells are not supported. | Automating report generation that requires individual cells for formula calculations after extracting merged cell values.
// AI Prompts: Generate C# code with Aspose.Cells to unmerge a specific range (e.g., B2:C2) and copy the original value to the right‑hand cell. | Show an Aspose.Cells snippet that detects a merged range, retrieves its top‑left value, unmerges the cells, and writes the value to an adjacent cell. | Explain how to preserve merged cell values when converting an Excel workbook to CSV using Aspose.Cells in .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace UnmergeAndCopyExample
{
    // Loads an Excel file (or creates a sample), reads the value from the merged range B2:C2, unmerges the cells, writes the original value into C2, and saves the result as a new workbook using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Ensure the input file exists; create a simple workbook if it does not.
                if (!File.Exists(inputPath))
                {
                    var sampleWb = new Workbook();
                    Worksheet sampleWs = sampleWb.Worksheets[0];
                    // Put a value in B2 and merge B2:C2 for demonstration.
                    sampleWs.Cells["B2"].PutValue("Sample");
                    sampleWs.Cells.CreateRange("B2", "C2").Merge();
                    sampleWb.Save(inputPath);
                }

                // Load the existing workbook.
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet.
                Worksheet worksheet = workbook.Worksheets[0];

                // Create a range that represents the merged cells B2:C2.
                Aspose.Cells.Range mergedRange = worksheet.Cells.CreateRange("B2", "C2");

                // Store the original value (the merged cell keeps its value in the top‑left cell B2).
                object originalValue = worksheet.Cells["B2"].Value;

                // Unmerge the range.
                mergedRange.UnMerge();

                // Copy the original value to an adjacent cell (C2 in this case).
                worksheet.Cells["C2"].PutValue(originalValue);

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook processed successfully. Output saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
