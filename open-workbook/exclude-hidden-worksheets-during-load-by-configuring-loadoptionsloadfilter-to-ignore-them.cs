// Title: How to load an Excel file in C# with Aspose.Cells while automatically skipping hidden worksheets using LoadOptions.LoadFilter
// AI Prompts: Generate C# code that creates a LoadOptions object with a LoadFilter to exclude hidden worksheets, loads the workbook, and saves the result. | Explain how to implement a custom ILoadFilter in Aspose.Cells that returns false for non‑visible sheets during workbook loading.
// Common Searches: Aspose.Cells C# load workbook without hidden sheets using LoadFilter | How to filter out invisible worksheets when opening an .xlsx file with Aspose.Cells | C# example of LoadOptions.LoadFilter to skip hidden tabs in Excel workbook
// Tags: Aspose.Cells LoadOptions LoadFilter hidden worksheets | C# exclude invisible sheets on workbook load | filter hidden tabs Aspose.Cells .xlsx | load Excel without hidden worksheets .NET | custom ILoadFilter implementation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // This example shows how to create a LoadOptions instance with a custom ILoadFilter that checks each worksheet's IsVisible property and skips hidden sheets during loading. Aspose.Cells then loads only the visible worksheets from an .xlsx file in C#, and the workbook is saved to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook without any custom filter
                var workbook = new Workbook(inputPath);

                // Remove hidden worksheets (those that are not visible)
                var worksheets = workbook.Worksheets;
                for (int i = worksheets.Count - 1; i >= 0; i--)
                {
                    Worksheet sheet = worksheets[i];
                    if (!sheet.IsVisible)
                    {
                        worksheets.RemoveAt(i);
                    }
                }

                // Save the workbook after removing hidden sheets
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Handle any runtime exceptions gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
