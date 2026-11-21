// Title: Hide selected pivot items in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells hide specific row field items in a pivot table | How to programmatically exclude confidential values from a pivot table with Aspose.Cells | Aspose.Cells hide pivot items based on a list of strings in .NET | Remove confidential rows from an Excel pivot table using Aspose.Cells API | Set PivotItem.IsHidden true for certain items in C# Aspose.Cells
// Tags: Aspose.Cells pivot item visibility C# | Excel pivot table confidential row filtering .NET | programmatic hiding of pivot items Aspose | C# hide pivot row field values | Aspose.Cells workbook pivot manipulation

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable, PivotField, PivotItem

// The example loads an existing Excel workbook, locates the first pivot table, and hides row‑field items whose values match a predefined confidential list, then saves the updated file.
class HidePivotItemsExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found on the worksheet.");
                return;
            }

            // Get the first pivot table (adjust index if there are multiple)
            PivotTable pivotTable = worksheet.PivotTables[0];

            // List of confidential items to hide
            List<string> confidentialItems = new List<string>
            {
                "Secret1",
                "Secret2",
                "ConfidentialData"
            };

            // Hide matching items in the first row field, if any row fields exist
            if (pivotTable.RowFields.Count > 0)
            {
                PivotField pivotField = pivotTable.RowFields[0]; // Change index to target another field

                foreach (PivotItem item in pivotField.PivotItems)
                {
                    // Use the Value property to get the item name
                    if (confidentialItems.Contains(item.Value))
                    {
                        item.IsHidden = true;
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
