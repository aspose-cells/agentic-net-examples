// Title: C# Aspose.Cells code to validate that every shape linked to a cell points to a valid, non‑empty cell
// AI Prompts: Write C# using Aspose.Cells that scans all shapes in each worksheet, checks the IsLinkedToCell and LinkToCell properties, and returns a list of shapes whose linked cell address is missing, invalid, or empty. | Create a reusable method in C# that validates shape‑to‑cell links across a workbook, verifies that the target cell contains data, and outputs detailed validation error messages. | Extend the validation routine to also flag shapes whose linked cell address lies outside the worksheet's used range and include the sheet name and shape name in the report.
// Common Searches: C# Aspose.Cells validate shape linked cell address is not empty | How to check if an Excel shape is linked to a valid cell using Aspose.Cells .NET | Detect shapes with invalid or empty linked cells in a workbook with Aspose.Cells | Aspose.Cells iterate shapes and verify linked cell data presence | Validate shape-to-cell links across all worksheets in C#
// Tags: shape linked cell validation Aspose.Cells | iterate worksheet shapes C# Aspose.Cells | reflection read shape link properties .NET | validate non‑empty linked cells Excel | shape link address existence check

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook with Aspose.Cells, iterates every worksheet and its shapes, uses reflection to read the IsLinkedToCell and LinkToCell properties, confirms that each linked address is syntactically valid and points to a cell containing a non‑empty value, collects any validation errors, reports them, and finally saves the workbook.
class ShapeCellLinkValidator
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file '{inputPath}' not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load workbook (lifecycle: load)
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        var validationErrors = new List<string>();

        // Iterate worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate shapes in the worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                bool isLinked = false;
                string linkedAddress = null;

                // Use reflection to access possible linking properties (covers different library versions)
                try
                {
                    var propIsLinked = shape.GetType().GetProperty("IsLinkedToCell");
                    if (propIsLinked != null)
                    {
                        isLinked = (bool)propIsLinked.GetValue(shape);
                    }

                    var propLink = shape.GetType().GetProperty("LinkToCell");
                    if (propLink != null)
                    {
                        linkedAddress = propLink.GetValue(shape) as string;
                    }
                }
                catch (Exception ex)
                {
                    // If reflection fails, treat as not linked
                    Console.WriteLine($"Warning: Unable to read link info for shape '{shape.Name}': {ex.Message}");
                    continue;
                }

                if (!isLinked || string.IsNullOrWhiteSpace(linkedAddress))
                {
                    // Skip shapes that are not linked or have no address
                    continue;
                }

                // Validate that the address refers to a non‑empty cell
                Cell linkedCell;
                try
                {
                    linkedCell = sheet.Cells[linkedAddress];
                }
                catch (Exception)
                {
                    validationErrors.Add(
                        $"Shape '{shape.Name}' on sheet '{sheet.Name}' has an invalid linked cell address '{linkedAddress}'.");
                    continue;
                }

                bool isEmpty = linkedCell.Value == null ||
                               (linkedCell.Value is string s && string.IsNullOrWhiteSpace(s));

                if (isEmpty)
                {
                    validationErrors.Add(
                        $"Shape '{shape.Name}' on sheet '{sheet.Name}' links to cell '{linkedAddress}' which is empty.");
                }
            }
        }

        // Output validation results
        if (validationErrors.Count == 0)
        {
            Console.WriteLine("All shape links reference valid, non‑empty cells.");
        }
        else
        {
            Console.WriteLine("Validation errors found:");
            foreach (string err in validationErrors)
            {
                Console.WriteLine(err);
            }
        }

        // Save workbook (lifecycle: save)
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
