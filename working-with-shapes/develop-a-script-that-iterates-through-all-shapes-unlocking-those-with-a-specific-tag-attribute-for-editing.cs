// Title: C# – Unlock Shapes by Name Tag in a Protected Worksheet with Aspose.Cells for .NET
// Description: Loads a workbook, protects the first worksheet while permitting drawing‑object edits, iterates all shapes, unlocks those whose Name matches a given tag, and saves the updated file.
// Keywords: Aspose.Cells | C# | .NET | unlock shape | shape tag | protected worksheet | worksheet.Shapes | IsLocked | drawing objects | Excel automation | code example | GitHub
// Common Searches: Aspose.Cells unlock shape by name | unlock shapes in protected Excel sheet C# | iterate worksheet shapes Aspose.Cells | shape.IsLocked false example | allow editing drawing objects Aspose.Cells | C# code unlock shapes with tag
// Developer Intent: Unlock only the shapes whose Name matches a specified tag so they remain editable on a protected sheet.
// Use Cases: Create a protected template where only callout shapes named "UnlockMe" stay editable for end users. | Automate a reporting workflow that unlocks placeholder shapes before inserting dynamic content. | Distribute a workbook that locks all drawing objects except those tagged for later modification.
// AI Prompts: Generate C# code using Aspose.Cells to unlock all shapes with a specific Name tag in a protected worksheet and save the workbook. | Show how to protect a worksheet, allow editing of drawing objects, then unlock selected shapes based on a custom tag. | Explain step‑by‑step how to loop through worksheet.Shapes, compare each shape's Name to a target tag, and set IsLocked = false.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads a workbook, protects the first worksheet while permitting drawing‑object edits, iterates all shapes, unlocks those whose Name matches a given tag, and saves the updated file.
class UnlockShapesByTag
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Allow editing of drawing objects even when the sheet is protected
        worksheet.Protection.AllowEditingObject = true;
        worksheet.Protect(ProtectionType.All);

        // Define the tag (here we use the shape's Name property as the tag)
        const string targetTag = "UnlockMe";

        // Iterate through all shapes in the worksheet
        foreach (Shape shape in worksheet.Shapes)
        {
            // Check if the shape's Name matches the target tag
            if (!string.IsNullOrEmpty(shape.Name) &&
                shape.Name.Equals(targetTag, StringComparison.OrdinalIgnoreCase))
            {
                // Unlock the shape for editing
                shape.IsLocked = false;
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
