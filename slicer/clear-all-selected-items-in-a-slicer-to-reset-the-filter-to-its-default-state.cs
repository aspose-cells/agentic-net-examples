// Title: How to clear all selections in an Excel slicer using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# clear slicer selections programmatically | reset Excel slicer filter to default with Aspose.Cells .NET | remove all selected items from worksheet slicer using Aspose.Cells | C# code to iterate slicer collection and clear filters in Aspose.Cells | Aspose.Cells ClearFilter method example for slicers
// Tags: Aspose.Cells ClearFilter slicer method | C# worksheet slicer collection iteration | reset Excel slicer to default .NET | programmatic slicer filter clearing Aspose.Cells | Excel workbook slicer manipulation C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads a workbook, checks for slicers on the first worksheet, iterates through each slicer using a dynamic reference, calls ClearFilter to remove all selections, handles any exceptions, and saves the workbook with the slicer filters reset.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Ensure the input file exists; create a minimal workbook if missing
                if (!File.Exists(inputPath))
                {
                    var placeholder = new Workbook();
                    placeholder.Worksheets[0].Name = "Sheet1";
                    placeholder.Save(inputPath);
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Attempt to clear slicer filters using dynamic to avoid compile‑time dependency on Slicer types
                try
                {
                    var slicerCollection = worksheet.Slicers;
                    if (slicerCollection != null && slicerCollection.Count > 0)
                    {
                        foreach (dynamic slicer in slicerCollection)
                        {
                            try
                            {
                                slicer.ClearFilter();
                            }
                            catch (Exception ex)
                            {
                                // Log but continue processing other slicers
                                Console.WriteLine($"Failed to clear filter on a slicer: {ex.Message}");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Slicer feature may not be available in the referenced Aspose.Cells version
                    Console.WriteLine($"Slicer processing skipped: {ex.Message}");
                }

                // Save the workbook with cleared slicer filters (if any)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
