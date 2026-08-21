// Title: C# – Update the LinkedCell of an Existing CheckBox Shape with Aspose.Cells
// Description: Loads a workbook, accesses the first worksheet, finds the CheckBox collection, changes the LinkedCell property of a selected CheckBox (e.g., to $C$5), and saves the file. Demonstrates how to keep a CheckBox synchronized with a new cell after layout changes using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# | update CheckBox LinkedCell | Excel CheckBox shape | change linked cell programmatically | .NET Excel checkbox example | modify CheckBox reference | Aspose.Cells shape manipulation | linked cell address update
// Common Searches: Aspose.Cells change linked cell of a CheckBox | C# update CheckBox LinkedCell in Excel file | how to set new linked cell for Excel CheckBox using Aspose | modify CheckBox shape reference after moving rows | Aspose.Cells example for updating CheckBox cell address
// Developer Intent: Modify the LinkedCell address of an existing CheckBox shape in an Excel workbook.
// Use Cases: Adjust a CheckBox after inserting or deleting rows/columns so it points to the correct data cell. | Re‑link CheckBox controls when redesigning a worksheet layout to maintain formula integrity. | Batch‑process multiple CheckBoxes in a template to match a new data schema.
// AI Prompts: Generate C# code that finds a CheckBox by its name in a worksheet and sets its LinkedCell to $D$10 using Aspose.Cells. | Create a method that iterates over all CheckBoxes on a sheet and updates each LinkedCell based on a dictionary of old‑to‑new addresses. | Provide C# error‑handling logic for cases where a CheckBox has no LinkedCell before attempting to update it with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads a workbook, accesses the first worksheet, finds the CheckBox collection, changes the LinkedCell property of a selected CheckBox (e.g., to $C$5), and saves the file. Demonstrates how to keep a CheckBox synchronized with a new cell after layout changes using Aspose.Cells for .NET.
class UpdateCheckBoxLinkedCell
{
    static void Main()
    {
        // Load an existing workbook that contains a CheckBox
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index or name as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Get the collection of CheckBox controls on the worksheet
        CheckBoxCollection checkBoxes = sheet.CheckBoxes;

        // Ensure there is at least one CheckBox
        if (checkBoxes.Count == 0)
        {
            Console.WriteLine("No CheckBox found on the worksheet.");
            return;
        }

        // Example: update the first CheckBox's linked cell
        // You can also locate a CheckBox by its name or other criteria
        CheckBox checkBox = checkBoxes[0];

        // New cell address that reflects the layout change
        string newLinkedCell = "$C$5";

        // Update the LinkedCell property
        checkBox.LinkedCell = newLinkedCell;

        // Optionally, verify the change
        Console.WriteLine($"CheckBox linked cell updated to: {checkBox.LinkedCell}");

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
