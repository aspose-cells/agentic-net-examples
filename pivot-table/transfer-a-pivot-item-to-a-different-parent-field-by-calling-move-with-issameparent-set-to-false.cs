// Title: Reparent a pivot table item to a different field using Aspose.Cells for .NET (C#)
// Common Searches: how to reparent a pivot table item to another field using Aspose.Cells C# | Aspose.Cells example for moving a pivot item to a different parent field | programmatically change pivot item parent in Excel with Aspose.Cells .NET | C# code to move a row field item to a column field in a pivot table using Aspose.Cells | move pivot item position without same parent Aspose.Cells API
// Tags: move pivot item Aspose.Cells C# | reparent pivot field item Aspose.Cells | pivot table item relocation .NET | Aspose.Cells change pivot item parent | Excel pivot item move C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExample
{
    // The sample loads an existing workbook, accesses the first pivot table, retrieves a specific row field item and a column field item by name, moves the source item to the target position with isSameParent set to false, refreshes and recalculates the pivot table, and saves the updated file.
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
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one pivot table
                if (sheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found in the worksheet.");
                    return;
                }

                // Get the first pivot table
                PivotTable pivotTable = sheet.PivotTables[0];

                // Identify the source field and item (replace with actual item name)
                if (pivotTable.RowFields.Count == 0)
                {
                    Console.WriteLine("No row fields in the pivot table.");
                    return;
                }
                PivotField sourceField = pivotTable.RowFields[0];
                PivotItem sourceItem;
                try
                {
                    sourceItem = sourceField.PivotItems["SourceItemName"]; // TODO: replace with real name
                }
                catch (Exception)
                {
                    Console.WriteLine("Source item not found.");
                    return;
                }

                // Identify the target field and item (replace with actual item name)
                if (pivotTable.ColumnFields.Count == 0)
                {
                    Console.WriteLine("No column fields in the pivot table.");
                    return;
                }
                PivotField targetField = pivotTable.ColumnFields[0];
                PivotItem targetItem;
                try
                {
                    targetItem = targetField.PivotItems["TargetItemName"]; // TODO: replace with real name
                }
                catch (Exception)
                {
                    Console.WriteLine("Target item not found.");
                    return;
                }

                // Move the source item to the target field position
                try
                {
                    sourceItem.Move(targetItem.Position, false);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error moving pivot item: {ex.Message}");
                    return;
                }

                // Refresh the pivot cache and recalculate the pivot table
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
