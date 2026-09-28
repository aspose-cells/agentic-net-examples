// Title: Group and lock Excel shapes by type using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that iterates every worksheet, gathers shapes by their Shape.Type, creates a group for each type that contains more than one shape, and sets the group's IsLocked property to true. | Update an existing Aspose.Cells workbook routine to add a method that groups shapes of the same type on a worksheet and locks the resulting group to prevent accidental modifications.
// Common Searches: Aspose.Cells C# group shapes of same type on a worksheet | lock grouped shapes in an Excel file using Aspose.Cells | prevent editing of shape groups with Aspose.Cells for .NET | example code to create shape groups and set IsLocked in Aspose.Cells | how to automatically lock shape groups in an Excel workbook via C#
// Tags: Aspose.Cells shape grouping C# | lock shape groups Aspose.Cells | group shapes by type .NET | prevent shape editing Aspose.Cells | Excel shape lock example

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeGrouping
{
    // // Loads an Excel workbook, iterates each worksheet, groups shapes that share the same Shape.Type into a single group, locks each group to avoid accidental edits, and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load workbook safely; create a new one if the input file does not exist
            try
            {
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                    workbook = new Workbook();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Process each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    GroupAndLockShapesByType(sheet);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing sheet \"{sheet.Name}\": {ex.Message}");
                }
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }

        /// <param name="sheet">The worksheet to process.</param>
        private static void GroupAndLockShapesByType(Worksheet sheet)
        {
            // Collect shapes by their numeric type identifier
            var shapesByType = new Dictionary<int, List<Shape>>();

            foreach (Shape shape in sheet.Shapes)
            {
                int typeKey = (int)shape.Type;

                if (!shapesByType.ContainsKey(typeKey))
                {
                    shapesByType[typeKey] = new List<Shape>();
                }

                shapesByType[typeKey].Add(shape);
            }

            // For each type with more than one shape, create a group and lock it
            foreach (KeyValuePair<int, List<Shape>> entry in shapesByType)
            {
                List<Shape> shapeList = entry.Value;

                if (shapeList.Count > 1)
                {
                    try
                    {
                        // Group the shapes; the Group method returns a new Shape representing the group
                        Shape groupShape = sheet.Shapes.Group(shapeList.ToArray());

                        // Lock the group to prevent accidental edits
                        groupShape.IsLocked = true;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to group shapes of type {entry.Key}: {ex.Message}");
                    }
                }
            }
        }
    }
}
