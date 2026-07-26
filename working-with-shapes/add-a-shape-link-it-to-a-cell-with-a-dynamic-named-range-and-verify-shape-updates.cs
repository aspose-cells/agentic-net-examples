// Title: Add ListBox Shape Linked to a Dynamic Named Range and Verify Cell Update – Aspose.Cells for .NET
// Description: This example creates a new workbook, fills column A with items, defines a dynamic named range "MyRange" using OFFSET and COUNTA, adds a ListBox shape, sets its input range to the dynamic range, links the selected item to cell B6, selects the second entry, updates the linked cell, reads the value for verification, and saves the file.
// Keywords: Aspose.Cells | C# ListBox shape | dynamic named range OFFSET | link shape to cell | verify linked cell | .NET spreadsheet | Excel dropdown shape | Aspose.Cells API | named range COUNTA | listbox selected index
// Common Searches: How to bind a ListBox shape to a dynamic named range using Aspose.Cells C# | Aspose.Cells C# example linking ListBox to a worksheet cell | Create expandable dropdown list in Excel with Aspose.Cells | Verify ListBox selected value updates linked cell Aspose.Cells | Dynamic range OFFSET formula Aspose.Cells .NET
// Developer Intent: Programmatically add a ListBox shape, bind its items to a dynamic named range, link its selection to a worksheet cell, and confirm the cell reflects the chosen item.
// Use Cases: Build interactive Excel reports where dropdown options grow with data entries. | Enable end‑users to select values from a shape‑based list that drives formulas. | Automate workbook generation with self‑updating data validation lists. | Create dashboards that synchronize shape selections with calculation cells.
// AI Prompts: Generate C# code with Aspose.Cells to insert a ListBox shape, set its input range to a dynamic OFFSET named range, link it to cell B6, select an item, and output the linked value. | Describe a testing approach to ensure a ListBox shape updates its linked cell after changing the selected index in Aspose.Cells for .NET. | List the steps to create a dynamic named range, attach it to a ListBox shape, and retrieve the selected value programmatically using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeDemo
{
    // This example creates a new workbook, fills column A with items, defines a dynamic named range "MyRange" using OFFSET and COUNTA, adds a ListBox shape, sets its input range to the dynamic range, links the selected item to cell B6, selects the second entry, updates the linked cell, reads the value for verification, and saves the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate column A with sample data
                sheet.Cells["A1"].PutValue("Item1");
                sheet.Cells["A2"].PutValue("Item2");
                sheet.Cells["A3"].PutValue("Item3");
                sheet.Cells["A4"].PutValue("Item4");

                // Define a dynamic named range "MyRange" using OFFSET and COUNTA
                // This range will expand automatically as items are added to column A
                int rangeIndex = workbook.Worksheets.Names.Add("MyRange");
                Name dynamicRange = workbook.Worksheets.Names[rangeIndex];
                dynamicRange.RefersTo = "=OFFSET($A$1,0,0,COUNTA($A:$A),1)";

                // Add a ListBox shape (dropdown) to the worksheet
                // Parameters: upper left row, upper left column, top, left, number of rows, number of columns
                ListBox listBox = (ListBox)sheet.Shapes.AddListBox(2, 2, 100, 100, 3, 20);

                // Link the ListBox to the dynamic named range for its items
                listBox.SetInputRange("MyRange", false, false);

                // Link the selected value of the ListBox to a specific cell (e.g., B6)
                listBox.SetLinkedCell("$B$6", false, false);

                // Set an initial selected index (0‑based). Here we select the second item ("Item2")
                listBox.SelectedIndex = 1;

                // Update the linked cell to reflect the selected value
                listBox.UpdateSelectedValue();

                // Verify: read the value from the linked cell and output it
                string linkedValue = sheet.Cells["B6"].StringValue;
                Console.WriteLine("Linked cell B6 value after update: " + linkedValue);

                // Save the workbook
                string outputPath = "ShapeLinkedToDynamicRange.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
