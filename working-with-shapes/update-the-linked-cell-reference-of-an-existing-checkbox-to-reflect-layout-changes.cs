// Title: Change the linked cell of an existing CheckBox shape in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Locate a CheckBox shape on a worksheet, assign a new address to its LinkedCell property, and save the workbook with Aspose.Cells in C#. | After rearranging a sheet layout, modify the linked cell reference of a checkbox by updating the CheckBox.LinkedCell value and persisting the changes.
// Common Searches: asp.net c# how to set linked cell for a checkbox shape using aspose.cells | update checkbox linked cell address programmatically in an existing Excel workbook | change linked cell of Excel form control checkbox with Aspose.Cells C# example | modify checkbox shape binding after rearranging cells in Aspose.Cells
// Tags: Aspose.Cells CheckBox LinkedCell property | C# edit Excel shape linked cell | programmatic checkbox binding Aspose.Cells | modify form control cell reference | update Excel checkbox link via API

using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an existing workbook, finds the first CheckBox shape on the first worksheet, sets its LinkedCell property to a new cell address (e.g., "D5"), and saves the updated file.
class Program
{
    static void Main()
    {
        // Load the existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the worksheet that contains the CheckBox (using index or name)
        Worksheet worksheet = workbook.Worksheets[0]; // or workbook.Worksheets["Sheet1"]

        // Locate the CheckBox shape.
        // If you know the name of the CheckBox, you can compare ws.Shapes[i].Name.
        CheckBox checkBox = null;
        for (int i = 0; i < worksheet.Shapes.Count; i++)
        {
            if (worksheet.Shapes[i] is CheckBox)
            {
                // Example: match by name
                // if (worksheet.Shapes[i].Name == "CheckBox1")
                // {
                //     checkBox = (CheckBox)worksheet.Shapes[i];
                //     break;
                // }

                // If name is not known, take the first CheckBox found
                checkBox = (CheckBox)worksheet.Shapes[i];
                break;
            }
        }

        if (checkBox != null)
        {
            // Update the linked cell reference to reflect the new layout.
            // Change "D5" to whatever cell you need the CheckBox to be linked to.
            checkBox.LinkedCell = "D5";
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
