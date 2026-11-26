// Title: Add a multi‑select ListBox shape populated with ten items to an Excel worksheet using Aspose.Cells for .NET
// Common Searches: how to create a ListBox shape with multi‑selection in an Excel file using Aspose.Cells C# | Aspose.Cells C# set ListBox InputRange to A1:A10 and link it to B1 | programmatically add a free‑floating ListBox to a worksheet with Aspose.Cells | enable multi‑select mode for a ListBox control in Aspose.Cells .NET example | populate ListBox items from a cell range in an Aspose.Cells workbook
// Tags: add ListBox shape Aspose.Cells | set ListBox multi‑selection Aspose.Cells | populate ListBox using InputRange Aspose.Cells | bind ListBox selection to worksheet cell Aspose.Cells | free‑floating ListBox placement Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ListBoxDemo
{
    // The example creates a new workbook, writes ten items (Item 1‑Item 10) into cells A1:A10, adds a ListBox shape at row 2 column 2, sets its InputRange to A1:A10, links the selected values to cell B1, enables multi‑selection, assigns a name and title, and saves the file as ListBoxMultiSelectDemo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Path where the resulting Excel file will be saved.
            const string outputPath = "ListBoxMultiSelectDemo.xlsx";

            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Access the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];

                // Populate ten entries in column A (A1:A10).
                for (int i = 0; i < 10; i++)
                {
                    // Cells are zero‑based for row/column indexes.
                    // Row i corresponds to cell A{i+1}.
                    sheet.Cells[i, 0].PutValue($"Item {i + 1}");
                }

                // Add a ListBox shape.
                // Parameters: upperLeftRow, top (pixel offset), upperLeftColumn, left (pixel offset), height (pixels), width (pixels)
                // Place it starting at row 2 (index 1) and column 2 (index 1) for visibility.
                ListBox listBox = sheet.Shapes.AddListBox(1, 0, 1, 0, 120, 200);

                // Configure the ListBox.
                listBox.InputRange = "A1:A10";          // Use the ten items as the source list.
                listBox.LinkedCell = "B1";              // Cell that will receive the selected value(s).
                listBox.SelectionType = SelectionType.Multi; // Enable multi‑selection mode.
                listBox.Placement = PlacementType.FreeFloating; // Optional: keep it free‑floating.

                // Optional: give the ListBox a name and a title.
                listBox.Name = "MultiSelectListBox";
                listBox.Title = "Select Items";

                // Save the workbook.
                workbook.Save(outputPath);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("ListBox created with multi‑selection enabled and ten entries populated.");
            }
            catch (Exception ex)
            {
                // Output any errors that occur during processing.
                Console.Error.WriteLine("An error occurred while creating the ListBox:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
