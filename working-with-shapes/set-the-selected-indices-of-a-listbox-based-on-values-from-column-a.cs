// Title: Set the SelectedIndex of an Aspose.Cells ListBox shape from values in column A using C#
// Common Searches: Aspose.Cells C# set ListBox selected item from worksheet range | programmatically select a ListBox entry by matching a cell value in Aspose.Cells | bind a ListBox to column A and set SelectedIndex with Aspose.Cells .NET | retrieve index of a target value in an Aspose.Cells ListBox shape | C# Aspose.Cells ListBox InputRange selection based on worksheet data
// Tags: Aspose.Cells ListBox SelectedIndex | Aspose.Cells ListBox InputRange binding | C# set ListBox selection from worksheet | Aspose.Cells populate ListBox from column A | Aspose.Cells workbook shape selection

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ListBoxSelectionDemo
{
    // The example creates a workbook, fills cells A2:A7 with department names, adds a ListBox shape bound to that range, searches for a target value (e.g., "MIS"), calculates its zero‑based position within the range, assigns that position to the ListBox's SelectedIndex property, and saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to save the generated workbook
            const string outputPath = "ListBoxSelectionDemo.xlsx";

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // 1. Populate column A with sample data (A2:A7)
            // ------------------------------------------------------------
            string[] sampleValues = { "Sales", "Finance", "MIS", "R&D", "Marketing", "HR" };
            for (int i = 0; i < sampleValues.Length; i++)
            {
                // Row index is 1‑based for Excel, so start at row 2 (index 1)
                cells[i + 1, 0].PutValue(sampleValues[i]); // column 0 = "A"
            }

            // ------------------------------------------------------------
            // 2. Add a ListBox shape and configure it
            // ------------------------------------------------------------
            // Parameters: topRow, top (pixel offset), leftColumn, left (pixel offset), height (px), width (px)
            ListBox listBox = sheet.Shapes.AddListBox(1, 0, 0, 0, 120, 200);
            listBox.Placement = PlacementType.FreeFloating;
            listBox.InputRange = "A2:A7";          // Bind the list box to the values we just added
            listBox.LinkedCell = "B1";             // Cell that will receive the selected value
            listBox.SelectionType = SelectionType.Single; // Single selection for this demo
            listBox.Title = "Department List";

            // ------------------------------------------------------------
            // 3. Determine which items should be selected based on column A
            // ------------------------------------------------------------
            // Example scenario: select the item whose text equals a target value.
            // In a real scenario the target values could come from another source.
            string targetValue = "MIS";

            // Find the zero‑based index of the target value within the input range.
            // The input range starts at row 2 (Excel row index 2, zero‑based 1).
            int startRow = 2; // Excel row number where the input range begins
            int selectedIndex = -1; // -1 means not found

            for (int i = 0; i < sampleValues.Length; i++)
            {
                if (string.Equals(sampleValues[i], targetValue, StringComparison.OrdinalIgnoreCase))
                {
                    selectedIndex = i; // zero‑based index relative to the input range
                    break;
                }
            }

            if (selectedIndex >= 0)
            {
                // Set the selected index on the ListBox.
                // SelectedIndex is zero‑based, matching the index we calculated.
                listBox.SelectedIndex = selectedIndex;
                Console.WriteLine($"Selected \"{targetValue}\" at index {selectedIndex}.");
            }
            else
            {
                Console.WriteLine($"Value \"{targetValue}\" not found in column A. No selection applied.");
            }

            // ------------------------------------------------------------
            // 4. Save the workbook
            // ------------------------------------------------------------
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error saving workbook: " + ex.Message);
            }
        }
    }
}
