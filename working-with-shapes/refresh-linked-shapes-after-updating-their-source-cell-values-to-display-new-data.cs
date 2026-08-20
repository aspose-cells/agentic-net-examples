// Title: Refresh a linked ListBox shape after changing its linked cell with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a ListBox shape linked to a cell, set its input range, change the linked cell value, and call UpdateSelectedValue to refresh the shape so the new selection appears, then saves the file.
// Keywords: Aspose.Cells | Refresh linked shape | UpdateSelectedValue | ListBox linked cell | C# Excel shape refresh | Aspose.Cells .NET | Linked shape synchronization
// Common Searches: how to refresh a ListBox shape after changing its linked cell in Aspose.Cells | Aspose.Cells update linked shape selection programmatically | C# refresh dropdown ListBox linked to a cell | UpdateSelectedValue method example Aspose.Cells | synchronize Excel form controls with cell values using Aspose.Cells
// Developer Intent: Refresh a ListBox shape so it reflects the new value set in its linked cell.
// Use Cases: Keep a dropdown ListBox in sync with a cell that is modified by code. | Refresh multiple form controls after batch updates to their source cells. | Build interactive Excel forms where programmatic cell changes automatically update linked shapes.
// AI Prompts: Show me C# code that refreshes a linked ListBox shape after updating its linked cell using Aspose.Cells. | Give an example of updating several linked shapes after changing their source cells in a workbook. | Explain the difference between UpdateSelectedValue and other refresh methods for linked shapes in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add a ListBox shape linked to a cell, set its input range, change the linked cell value, and call UpdateSelectedValue to refresh the shape so the new selection appears, then saves the file.
    public class RefreshLinkedShapesDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate source data for the ListBox input range
                sheet.Cells["A1"].Value = "Option 1";
                sheet.Cells["A2"].Value = "Option 2";
                sheet.Cells["A3"].Value = "Option 3";

                // Add a ListBox shape (dropdown) to the worksheet
                // Parameters: upper left row, upper left column, top, left, width, height
                Shape listBoxShape = sheet.Shapes.AddListBox(2, 0, 2, 0, 130, 130);

                // Set the range that provides the list items
                listBoxShape.SetInputRange("$A$1:$A$3", false, false);

                // Link the selected value of the ListBox to cell B1
                listBoxShape.SetLinkedCell("$B$1", false, true);

                // Initial selection: set linked cell to the second option (index 2)
                sheet.Cells["B1"].Value = 2; // ListBox uses 1‑based index for selection
                // Refresh the shape so it reflects the linked cell value
                listBoxShape.UpdateSelectedValue();

                // Verify the selection (optional)
                ListBox listBox = (ListBox)listBoxShape;
                Console.WriteLine("Initially selected: " + (listBox.IsSelected(1) ? "Option 2" : "None"));

                // Change the linked cell value to select a different option
                sheet.Cells["B1"].Value = 3; // Select "Option 3"
                // Refresh the shape again to display the new selection
                listBoxShape.UpdateSelectedValue();

                // Verify the new selection
                Console.WriteLine("After update selected: " + (listBox.IsSelected(2) ? "Option 3" : "None"));

                // Save the workbook to a file
                workbook.Save("RefreshLinkedShapesDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            RefreshLinkedShapesDemo.Run();
        }
    }
}
