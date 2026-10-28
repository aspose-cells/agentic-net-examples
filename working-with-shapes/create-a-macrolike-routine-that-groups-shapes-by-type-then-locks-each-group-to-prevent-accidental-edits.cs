// Title: C# Aspose.Cells example: Group Excel shapes by type and lock the groups to prevent editing
// Common Searches: how to programmatically group shapes of the same type in an Excel file using Aspose.Cells for .NET | Aspose.Cells C# lock shape groups to avoid accidental changes | using reflection to invoke AddGroup on worksheet shapes with Aspose.Cells | C# code to lock multiple Excel shapes after grouping them by type | prevent editing of grouped shapes in a workbook with Aspose.Cells
// Tags: group shapes by type Aspose.Cells | lock shape groups Excel .NET | AddGroup method reflection Aspose.Cells | prevent accidental shape edits C# | shape collection processing Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an Excel workbook, iterates each worksheet, groups shapes that share the same Type, uses reflection to call the AddGroup method (when available) to create a shape group, adds the matching shapes to the group, locks the group and assigns a name, falls back to locking each shape individually if grouping is unsupported, and finally saves the modified workbook.
class ShapeGroupingAndLocking
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Process each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Group shapes by their type name
                Dictionary<string, List<Shape>> shapesByType = new Dictionary<string, List<Shape>>(StringComparer.OrdinalIgnoreCase);

                foreach (Shape shape in sheet.Shapes)
                {
                    string typeKey = shape.Type.ToString();
                    if (!shapesByType.ContainsKey(typeKey))
                        shapesByType[typeKey] = new List<Shape>();
                    shapesByType[typeKey].Add(shape);
                }

                // For each type that has more than one shape, attempt to create a group
                foreach (var entry in shapesByType)
                {
                    List<Shape> shapeList = entry.Value;
                    if (shapeList.Count < 2)
                        continue; // No need to group a single shape

                    // Try to locate the AddGroup method via reflection (covers different API versions)
                    MethodInfo addGroupMethod = sheet.Shapes.GetType().GetMethod("AddGroup", BindingFlags.Instance | BindingFlags.Public);
                    if (addGroupMethod != null)
                    {
                        // Create the group shape
                        object groupObj = addGroupMethod.Invoke(sheet.Shapes, null);
                        dynamic group = groupObj; // Use dynamic to avoid compile‑time type dependency

                        // Add each shape to the group
                        foreach (Shape s in shapeList)
                        {
                            group.Add(s);
                        }

                        // Lock the group and assign a name
                        group.IsLocked = true;
                        group.Name = $"Group_{entry.Key}";
                    }
                    else
                    {
                        // Fallback: lock each shape individually if grouping is unavailable
                        foreach (Shape s in shapeList)
                        {
                            s.IsLocked = true;
                        }
                    }
                }
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
