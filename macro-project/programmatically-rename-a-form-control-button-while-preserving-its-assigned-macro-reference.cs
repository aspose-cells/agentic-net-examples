// Title: Rename an Excel Form Control Button and Preserve Its Macro Using Aspose.Cells for .NET (C#)
// Description: Loads a workbook, finds a Button shape named "OldButtonName" on the first worksheet, saves its MacroName, changes the button's Name to "NewButtonName", re‑applies the saved MacroName to keep the macro link intact, and saves the result as a new file.
// Keywords: Aspose.Cells | C# | .NET | Excel form control | button rename | MacroName | preserve macro | Button shape | programmatic rename | Excel macro link
// Common Searches: Rename Excel form control button with Aspose.Cells C# | Keep macro reference when changing button name in .NET | Update Button.Name without losing MacroName using Aspose.Cells | Find and rename button shape in workbook programmatically
// Developer Intent: Rename a form control button in an Excel workbook while keeping its assigned macro unchanged.
// Use Cases: Standardize button identifiers across generated workbooks without breaking existing macros. | Refactor legacy Excel templates by renaming controls while maintaining macro connections. | Automate bulk updates of button names in multiple worksheets, ensuring macro functionality remains intact.
// AI Prompts: Generate C# code with Aspose.Cells that locates a button named 'SubmitBtn', renames it to 'SendBtn', and retains its MacroName. | Explain how to iterate through worksheet shapes to find a specific Button and modify its Name property without affecting the linked macro.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads a workbook, finds a Button shape named "OldButtonName" on the first worksheet, saves its MacroName, changes the button's Name to "NewButtonName", re‑applies the saved MacroName to keep the macro link intact, and saves the result as a new file.
class RenameButtonDemo
{
    static void Main()
    {
        // Load an existing workbook that contains a form control button
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Locate the button by its current name
        Button button = null;
        foreach (Shape shape in worksheet.Shapes)
        {
            if (shape is Button && shape.Name == "OldButtonName")
            {
                button = (Button)shape;
                break;
            }
        }

        if (button != null)
        {
            // Store the existing macro reference
            string existingMacro = button.MacroName;

            // Rename the button while keeping the macro unchanged
            button.Name = "NewButtonName";

            // Reassign the macro name to guarantee it is preserved
            button.MacroName = existingMacro;
        }
        else
        {
            Console.WriteLine("Button with the specified name was not found.");
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
