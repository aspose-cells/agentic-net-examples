// Title: C# – Unlock Shapes by Tag in an Excel Workbook with Aspose.Cells
// Description: Load a workbook, enable object editing on protected sheets, iterate through all shapes, identify those whose Name (or AlternativeText) contains a given tag, set IsLocked = false, and save the updated file.
// Keywords: Aspose.Cells C# unlock shape | shape tag Aspose.Cells | iterate shapes workbook | IsLocked false Aspose.Cells | worksheet protection edit objects | Excel shape unlocking C# | tag‑based shape selection | Aspose.Cells shape collection | batch unlock Excel shapes | C# Excel automation Aspose
// Common Searches: how to unlock shapes with a specific tag using Aspose.Cells C# | iterate all shapes in an Excel workbook and set IsLocked false | enable editing of objects on a protected worksheet Aspose.Cells | unlock Excel shapes by name keyword in .NET | Aspose.Cells example for unlocking selected shapes
// Developer Intent: Programmatically remove the lock from shapes that match a predefined tag so they remain editable even when the worksheet is protected.
// Use Cases: Prepare a template where only diagram elements marked "UnlockMe" can be edited by end users. | Automate bulk processing of multiple worksheets to expose specific shapes while keeping the rest locked. | Create a reporting workbook that protects the sheet but allows users to modify only tagged comment boxes or icons.
// AI Prompts: Write C# code with Aspose.Cells that unlocks all shapes whose Name contains "UnlockMe" and then re‑protects each worksheet. | Show how to store a custom identifier in a shape's AlternativeText property and unlock shapes based on that identifier using Aspose.Cells. | Explain handling of shape unlocking when the workbook contains shared formulas and several protected sheets.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeUnlocker
{
    // Load a workbook, enable object editing on protected sheets, iterate through all shapes, identify those whose Name (or AlternativeText) contains a given tag, set IsLocked = false, and save the updated file.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Define the tag that identifies shapes to be unlocked
            const string targetTag = "UnlockMe";

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Ensure that drawing objects can be edited when the sheet is protected
                sheet.Protection.AllowEditingObject = true;

                // Iterate through each shape in the current worksheet
                ShapeCollection shapes = sheet.Shapes;
                for (int i = 0; i < shapes.Count; i++)
                {
                    Shape shape = shapes[i];

                    // Check if the shape's Name (used here as a tag) matches the target tag
                    // Adjust this condition if you store the tag in a different property
                    if (shape.Name != null && shape.Name.Contains(targetTag, StringComparison.OrdinalIgnoreCase))
                    {
                        // Unlock the shape so it can be edited even when the sheet is protected
                        shape.IsLocked = false;
                    }
                }

                // (Optional) Protect the worksheet after unlocking the desired shapes
                // sheet.Protect(ProtectionType.All);
            }

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}
