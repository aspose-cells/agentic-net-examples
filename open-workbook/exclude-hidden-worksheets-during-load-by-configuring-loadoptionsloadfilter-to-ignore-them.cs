// Title: Load an Excel workbook in C# with Aspose.Cells and delete hidden worksheets before saving
// Common Searches: how to delete hidden worksheets after loading an Excel file with Aspose.Cells in C# | Aspose.Cells C# load workbook and remove invisible sheets programmatically | skip hidden tabs when processing Excel workbooks using Aspose.Cells .NET | remove hidden worksheets from a loaded workbook before saving with Aspose.Cells | C# code to filter out hidden worksheets after opening an Excel file with Aspose.Cells
// Tags: Aspose.Cells remove hidden worksheets C# | load Excel workbook without hidden sheets Aspose.Cells | delete invisible worksheets after workbook load .NET | filter hidden worksheets on load Aspose.Cells API | exclude hidden tabs when saving Excel file C# | Aspose.Cells workbook cleanup hidden sheets

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel file using Aspose.Cells, iterates through the worksheets to collect indices of those that are not visible, removes the hidden worksheets from the workbook, and then saves the cleaned workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file.
            string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook without a custom filter.
                Workbook workbook = new Workbook(inputPath);

                // Remove hidden worksheets after loading.
                // Collect indices of hidden sheets first to avoid modifying the collection while iterating.
                var hiddenSheetIndices = new System.Collections.Generic.List<int>();
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    Worksheet sheet = workbook.Worksheets[i];
                    if (!sheet.IsVisible)
                    {
                        hiddenSheetIndices.Add(i);
                    }
                }

                // Delete hidden sheets in reverse order to keep indices valid.
                for (int i = hiddenSheetIndices.Count - 1; i >= 0; i--)
                {
                    workbook.Worksheets.RemoveAt(hiddenSheetIndices[i]);
                }

                // (Optional) Perform additional operations on the loaded workbook here.

                // Save the workbook to a new file.
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime exceptions gracefully.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
