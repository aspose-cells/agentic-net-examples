// Title: Ensure Unique Shape.Name for ActiveX Controls in Aspose.Cells (C#)
// Description: This C# example creates a workbook, adds ActiveX controls with duplicate shape names, and runs a routine that scans every worksheet, collects existing names case‑insensitively, and renames conflicts by appending a numeric suffix. The workbook is saved with each ActiveX control having a distinct Shape.Name, preventing reference errors.
// Keywords: Aspose.Cells ActiveX unique name | C# duplicate ActiveX shape name | validate ActiveX control names | rename Excel ActiveX controls | Shape.Name uniqueness Aspose
// Common Searches: how to make ActiveX control names unique in Aspose.Cells | duplicate ActiveX shape names .NET | ensure distinct Shape.Name for Excel ActiveX controls | Aspose.Cells rename duplicate ActiveX controls | C# validate ActiveX control names before saving
// Developer Intent: Guarantee that every ActiveX control in an Aspose.Cells workbook has a unique Shape.Name prior to saving the file.
// Use Cases: Automated report generation where programmatically added ActiveX controls must be uniquely addressable. | Pre‑deployment validation of workbooks containing multiple ActiveX controls across sheets. | Refactoring legacy spreadsheets that have overlapping control identifiers.
// AI Prompts: Generate a C# method for Aspose.Cells that iterates all worksheets and makes ActiveX Shape.Name values unique by adding a suffix. | Write code that throws an exception when duplicate ActiveX control names are detected in an Aspose.Cells workbook. | Create an xUnit test that confirms the EnsureUniqueActiveXNames function correctly renames duplicate control names.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

// This C# example creates a workbook, adds ActiveX controls with duplicate shape names, and runs a routine that scans every worksheet, collects existing names case‑insensitively, and renames conflicts by appending a numeric suffix. The workbook is saved with each ActiveX control having a distinct Shape.Name, preventing reference errors.
class ActiveXUniqueNameValidator
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add some ActiveX controls with intentional duplicate names
        AddActiveXControl(sheet, ControlType.CheckBox, "MyControl", 1, 1);
        AddActiveXControl(sheet, ControlType.CommandButton, "MyControl", 3, 1);
        AddActiveXControl(sheet, ControlType.ComboBox, "AnotherControl", 5, 1);

        // Validate and make ActiveX control names unique before saving
        EnsureUniqueActiveXNames(workbook);

        // Save the workbook
        workbook.Save("UniqueActiveXNames.xlsx");
    }

    // Helper method to add an ActiveX control and assign an initial shape name
    private static void AddActiveXControl(Worksheet sheet, ControlType type, string shapeName, int row, int column)
    {
        // Add the ActiveX control to the worksheet
        Shape shape = sheet.Shapes.AddActiveXControl(type, row, 0, column, 0, 100, 30);
        // Set the shape's name (may be duplicate)
        shape.Name = shapeName;

        // Set a simple property so the control is functional
        if (type == ControlType.CheckBox)
        {
            ((CheckBoxActiveXControl)shape.ActiveXControl).Caption = "Check";
        }
        else if (type == ControlType.CommandButton)
        {
            shape.ActiveXControl.IsEnabled = true;
        }
    }

    // Ensures each ActiveX control's Shape.Name is unique across the entire workbook
    private static void EnsureUniqueActiveXNames(Workbook workbook)
    {
        // Track names that have already been used
        HashSet<string> existingNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (Worksheet ws in workbook.Worksheets)
        {
            foreach (Shape shape in ws.Shapes)
            {
                // Process only shapes that host an ActiveX control
                if (shape.ActiveXControl != null)
                {
                    string originalName = shape.Name;
                    // If the shape has no name, start with a default base name
                    if (string.IsNullOrEmpty(originalName))
                    {
                        originalName = "ActiveXControl";
                    }

                    string uniqueName = originalName;
                    int suffix = 1;
                    // Append a numeric suffix until the name becomes unique
                    while (!existingNames.Add(uniqueName))
                    {
                        uniqueName = $"{originalName}_{suffix}";
                        suffix++;
                    }

                    // Assign the unique name back to the shape
                    shape.Name = uniqueName;
                }
            }
        }
    }
}
