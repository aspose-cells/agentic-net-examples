// Title: Update a ListBox shape after changing its linked cell with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, populate cells A1:A5, add a ListBox shape, set its input range, link it to cell B1, modify the linked cell value, call UpdateSelectedValue to refresh the shape's selection, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# ListBox shape | UpdateSelectedValue | linked cell refresh | SetLinkedCell | SetInputRange | shape synchronization | Aspose.Cells .NET example | programmatic ListBox selection
// Common Searches: Aspose.Cells refresh ListBox after linked cell change | UpdateSelectedValue usage in Aspose.Cells | How to sync ListBox shape with cell value in C# | SetLinkedCell and UpdateSelectedValue example | Aspose.Cells shape linked cell update
// Developer Intent: Synchronize a ListBox shape's selected item with a new value in its linked cell using Aspose.Cells for .NET.
// Use Cases: Select a ListBox item based on a calculation or formula result. | Keep form controls (ListBox, ComboBox) in sync with dynamic worksheet data. | Programmatically verify that a shape reflects the current linked cell value after updates.
// AI Prompts: Write C# code that changes the linked cell of a ListBox shape and calls UpdateSelectedValue to refresh the selection with Aspose.Cells. | Explain when and why UpdateSelectedValue must be invoked after modifying a linked cell in Aspose.Cells. | Provide an example of linking a ComboBox shape to a cell and updating its selected value using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, populate cells A1:A5, add a ListBox shape, set its input range, link it to cell B1, modify the linked cell value, call UpdateSelectedValue to refresh the shape's selection, and save the file using Aspose.Cells for .NET.
    public class UpdateLinkedShapeDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate source data for the ListBox (A1:A5)
                for (int i = 0; i < 5; i++)
                {
                    sheet.Cells[i, 0].Value = $"Item {i + 1}";
                }

                // Add a ListBox shape to the worksheet
                // Parameters: upperRow, upperColumn, lowerRow, lowerColumn, width, height
                ListBox listBoxShape = sheet.Shapes.AddListBox(6, 0, 6, 0, 120, 120);

                // Set the range that provides the list items
                listBoxShape.SetInputRange("$A$1:$A$5", false, false);

                // Link the ListBox selected value to cell B1
                listBoxShape.SetLinkedCell("$B$1", false, true);

                // Initially set the linked cell value to the first item (index 0)
                sheet.Cells["B1"].Value = 0;
                // Update the shape so it reflects the linked cell value
                listBoxShape.UpdateSelectedValue();

                // Change the linked cell value to select the third item (index 2)
                sheet.Cells["B1"].Value = 2;

                // After changing the source cell, refresh the shape selection
                listBoxShape.UpdateSelectedValue();

                // Optional: verify the selection programmatically
                if (listBoxShape.IsSelected(2))
                {
                    Console.WriteLine("Third option is now selected after updating the linked cell.");
                }

                // Save the workbook
                string outputPath = "UpdateLinkedShapeDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            UpdateLinkedShapeDemo.Run();
        }
    }
}
