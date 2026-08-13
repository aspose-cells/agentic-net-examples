// Title: Validate Unique ActiveX Control Names in Aspose.Cells (C#) Before Saving
// Description: C# example that creates a workbook, adds ActiveX checkboxes, then scans all worksheets to ensure each ActiveX control has a distinct Name. Empty names get a default, duplicates are resolved with a numeric suffix using a case‑insensitive HashSet, and the file is saved.
// Keywords: Aspose.Cells | C# | ActiveX control naming | duplicate shape names | unique name validation | Excel workbook | HashSet | case‑insensitive | shape renaming | GitHub example
// Common Searches: Aspose.Cells ensure unique ActiveX control names | C# rename duplicate ActiveX shapes in Excel | how to prevent ActiveX name collisions with Aspose | validate shape names before saving workbook | case insensitive ActiveX name check Aspose.Cells
// Developer Intent: Guarantee that every ActiveX control in an Aspose.Cells workbook has a unique Name property prior to saving.
// Use Cases: Detect and rename duplicate ActiveX control names across all worksheets. | Assign a default name based on the control type when the Name property is empty. | Automatically append a numeric suffix to conflicting names to maintain uniqueness.
// AI Prompts: Generate a C# method for Aspose.Cells that enforces unique Name values on all ActiveX controls, adding numeric suffixes for duplicates. | Show code that assigns default names to ActiveX controls with blank names and resolves naming conflicts using a case‑insensitive HashSet. | Explain how to extend the EnsureUniqueActiveXControlNames routine to support custom naming patterns and locale‑specific case rules.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

namespace AsposeCellsActiveXValidation
{
    // C# example that creates a workbook, adds ActiveX checkboxes, then scans all worksheets to ensure each ActiveX control has a distinct Name. Empty names get a default, duplicates are resolved with a numeric suffix using a case‑insensitive HashSet, and the file is saved.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add two ActiveX controls with the same default name to demonstrate validation
            Shape shape1 = sheet.Shapes.AddActiveXControl(ControlType.CheckBox, 2, 0, 2, 0, 100, 30);
            Shape shape2 = sheet.Shapes.AddActiveXControl(ControlType.CheckBox, 5, 0, 5, 0, 100, 30);

            // Both shapes receive the same default name ("CheckBox 1") – we will fix this
            Console.WriteLine($"Before validation: Shape1.Name = {shape1.Name}, Shape2.Name = {shape2.Name}");

            // Validate and ensure unique names for all ActiveX controls
            EnsureUniqueActiveXControlNames(workbook);

            // After validation, duplicate names are resolved
            Console.WriteLine($"After validation: Shape1.Name = {shape1.Name}, Shape2.Name = {shape2.Name}");

            // Save the workbook (using the standard Save method as required by lifecycle rules)
            workbook.Save("ValidatedActiveXControls.xlsx");
        }

        /// <param name="workbook">The workbook to validate.</param>
        static void EnsureUniqueActiveXControlNames(Workbook workbook)
        {
            // Keep track of names that have already been used
            HashSet<string> usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // Iterate through all worksheets
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Iterate through all shapes in the worksheet
                foreach (Shape shape in ws.Shapes)
                {
                    // Only process shapes that host an ActiveX control
                    if (shape.ActiveXControl != null)
                    {
                        string originalName = shape.Name;

                        // If the name is empty, assign a default based on control type
                        if (string.IsNullOrWhiteSpace(originalName))
                        {
                            originalName = shape.ActiveXControl.Type.ToString();
                            shape.Name = originalName;
                        }

                        string uniqueName = originalName;
                        int suffix = 1;

                        // Resolve duplicates by appending a numeric suffix
                        while (usedNames.Contains(uniqueName))
                        {
                            uniqueName = $"{originalName}_{suffix}";
                            suffix++;
                        }

                        // Update the shape's name if it was changed
                        if (!uniqueName.Equals(shape.Name, StringComparison.OrdinalIgnoreCase))
                        {
                            shape.Name = uniqueName;
                        }

                        // Record the name as used
                        usedNames.Add(shape.Name);
                    }
                }
            }
        }
    }
}
