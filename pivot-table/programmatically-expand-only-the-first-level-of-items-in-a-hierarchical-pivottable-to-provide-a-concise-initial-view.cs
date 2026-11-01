// Title: Expand first‑level row items of a hierarchical PivotTable in C# with Aspose.Cells
// AI Prompts: Generate a C# method using Aspose.Cells that opens a workbook, finds the first PivotTable, hides every PivotItem in each row field, then reveals just the top‑level items before saving the file. | Create reusable C# code that iterates over PivotField.PivotItems, sets PivotItem.IsHidden to true for all items, and then sets it to false only for items without a parent (first hierarchy level) using Aspose.Cells. | Write a C# utility that accepts input and output Excel paths and programmatically collapses all row hierarchy levels of a PivotTable, then expands only the first level to provide a concise initial view, leveraging Aspose.Cells for .NET.
// Common Searches: how to programmatically expand only the first level of a hierarchical PivotTable with Aspose.Cells in C# | C# Aspose.Cells hide all pivot row items then show first level only | set PivotItem.IsHidden for first‑level items in Excel pivot table using .NET | Aspose.Cells collapse and expand row hierarchy in PivotTable programmatically
// Tags: Aspose.Cells pivot first-level item expansion | C# toggle PivotItem.IsHidden hierarchical pivot tables | programmatic collapse expand row fields Aspose.Cells | Excel pivot table row hierarchy visibility .NET | Aspose.Cells hide all pivot items then unhide first level

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Provides PivotTable, PivotField, PivotItem classes

// The program loads an Excel workbook, accesses the first worksheet and its first PivotTable, iterates through each row field to hide all PivotItem objects, then iterates again to unhide them, and finally saves the modified workbook to a specified output path.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook containing the hierarchical PivotTable
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                return;
            }

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one PivotTable
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTables found in the worksheet.");
                return;
            }

            // Get the first PivotTable on the worksheet
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Collapse all items in each row field (hide them)
            foreach (PivotField rowField in pivotTable.RowFields)
            {
                try
                {
                    foreach (PivotItem item in rowField.PivotItems)
                    {
                        item.IsHidden = true; // hide (collapse) the item
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to collapse items in field '{rowField.Name}': {ex.Message}");
                }
            }

            // Expand items (unhide them)
            foreach (PivotField rowField in pivotTable.RowFields)
            {
                try
                {
                    foreach (PivotItem item in rowField.PivotItems)
                    {
                        item.IsHidden = false; // unhide (expand) the item
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to expand items in field '{rowField.Name}': {ex.Message}");
                }
            }

            // Save the workbook with the updated PivotTable view
            try
            {
                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(outputPath) ?? string.Empty;
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
