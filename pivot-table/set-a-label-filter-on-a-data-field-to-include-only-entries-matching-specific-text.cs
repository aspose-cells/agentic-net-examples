// Title: How to hide pivot table row items that don't contain a specific text using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide pivot row items that do not contain a keyword | set label filter on a pivot table field programmatically with Aspose.Cells | filter pivot table items by text in .NET using Aspose.Cells API | apply custom label filter to Excel pivot table rows using C# | refresh pivot table after modifying item visibility with Aspose.Cells
// Tags: Aspose.Cells pivot label filter C# | hide pivot items by text Aspose.Cells | programmatic pivot item visibility .NET | refresh pivot table after filter Aspose.Cells | Excel pivot table filtering using C# API

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The example loads an existing Excel workbook, accesses its first worksheet and pivot table, selects a row field, and hides any pivot items whose text does not contain the specified keyword. After updating item visibility, the pivot table is refreshed and recalculated, and the workbook is saved to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string filterKeyword = "Target";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Get the first pivot table in the worksheet
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found in the worksheet.");
                    return;
                }

                PivotTable pivotTable = worksheet.PivotTables[0];

                // Choose a pivot field to apply the label filter.
                // Here we use the first row field; modify the index as required.
                if (pivotTable.RowFields.Count == 0)
                {
                    Console.WriteLine("The pivot table does not contain any row fields to filter.");
                    return;
                }

                PivotField pivotField = pivotTable.RowFields[0];

                // Apply a simple label filter: hide items that do NOT contain the keyword.
                foreach (PivotItem item in pivotField.PivotItems)
                {
                    // Ensure the item value is not null before checking
                    string itemText = item.Value?.ToString() ?? string.Empty;
                    if (!itemText.Contains(filterKeyword, StringComparison.OrdinalIgnoreCase))
                    {
                        item.IsHidden = true; // Exclude this item from the pivot view
                    }
                }

                // Refresh the pivot table to reflect the changes
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
