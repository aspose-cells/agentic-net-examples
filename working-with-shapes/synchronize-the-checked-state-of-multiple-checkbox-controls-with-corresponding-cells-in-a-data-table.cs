// Title: C# Example: Synchronize Aspose.Cells CheckBox Shapes with Linked Worksheet Cells
// Description: Demonstrates how to create a workbook, add multiple CheckBox shapes, link each CheckBox to a specific cell, set alternating checked states, save the file, reload it, read the linked cell values to confirm synchronization, and finally persist any updates using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# checkbox example | link CheckBox to cell Aspose.Cells | synchronize checkbox state Excel | Aspose.Cells CheckBox shapes | read checkbox values after save | programmatically set checkbox state | Excel checklist with Aspose.Cells | C# Aspose.Cells workbook lifecycle
// Common Searches: How to bind an Aspose.Cells CheckBox to a worksheet cell in C# | Read linked cell value of a CheckBox after loading an Excel file | Set initial checked state for multiple checkboxes with Aspose.Cells | Aspose.Cells example for syncing checkbox state with cells | C# code to add and link CheckBox shapes in Excel
// Developer Intent: Programmatically keep the checked state of several CheckBox controls in sync with corresponding worksheet cells using Aspose.Cells for .NET.
// Use Cases: Create a printable checklist where each item has a hidden checkbox linked to a Boolean cell that records user selections. | Load an existing Excel file, read the Boolean values of linked cells to determine which items were marked, and act on the results in C#. | Update the checked state of checkboxes by modifying the linked cell values based on external data, then save the workbook.
// AI Prompts: Generate C# code with Aspose.Cells that adds CheckBox shapes, links each to a specific cell, and alternates the initial checked state. | Show how to open a workbook containing linked CheckBox shapes, iterate through the CheckBox collection, and output each linked cell's Boolean value. | Provide an example that changes the linked cell values for a set of checkboxes based on a data source and saves the updated workbook using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, add multiple CheckBox shapes, link each CheckBox to a specific cell, set alternating checked states, save the file, reload it, read the linked cell values to confirm synchronization, and finally persist any updates using Aspose.Cells for .NET.
class SyncCheckBoxes
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add headers
        sheet.Cells["A1"].PutValue("Item");
        sheet.Cells["C1"].PutValue("Checked");

        // Add items and checkboxes for rows 2 to 6
        for (int row = 2; row <= 6; row++)
        {
            // Put item name in column A
            sheet.Cells[row - 1, 0].PutValue("Item " + (row - 1));

            // Add a checkbox at column B (row index, column index, height, width)
            int cbIndex = sheet.CheckBoxes.Add(row - 1, 1, 20, 100);
            CheckBox checkBox = sheet.CheckBoxes[cbIndex];
            checkBox.Text = ""; // No visible text

            // Link the checkbox to column C of the same row
            string linkedCell = $"C{row}";
            checkBox.LinkedCell = linkedCell;

            // Set an initial checked state (alternating true/false)
            checkBox.Value = (row % 2 == 0);
        }

        // Save the workbook (create lifecycle)
        workbook.Save("CheckBoxSyncDemo.xlsx");

        // Load the workbook to verify synchronization (load lifecycle)
        Workbook loadedWorkbook = new Workbook("CheckBoxSyncDemo.xlsx");
        Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

        // Output the linked cell values for each checkbox
        for (int i = 0; i < loadedSheet.CheckBoxes.Count; i++)
        {
            CheckBox cb = loadedSheet.CheckBoxes[i];
            string linkedCell = cb.LinkedCell;
            bool cellValue = loadedSheet.Cells[linkedCell].BoolValue;
            Console.WriteLine($"{cb.Name} linked to {linkedCell} = {cellValue}");
        }

        // Save any changes (save lifecycle)
        loadedWorkbook.Save("CheckBoxSyncDemo_Updated.xlsx");
    }
}
