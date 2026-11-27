// Title: Add a ComboBox shape with five items and custom drop‑down width in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells create ComboBox shape and set dropdown list width | populate ComboBox items from a cell range and define drop‑down width with Aspose.Cells | Aspose.Cells add ComboBox to worksheet and link selected value to a cell in C# | set visible lines and dropdown width for a ComboBox shape using Aspose.Cells .NET | create Excel ComboBox with custom list width using Aspose.Cells API
// Tags: Aspose.Cells ComboBox shape creation | ComboBox input range Aspose.Cells | ComboBox dropdown width configuration | Link ComboBox to cell Aspose.Cells | Set ComboBox visible lines Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsComboBoxDemo
{
    // The example creates a new workbook, writes five strings into cells A1:A5, adds a ComboBox shape at cell C2, links its selected value to B1, sets the input range to A1:A5, configures the drop‑down list width to 200 pixels, specifies five visible lines, and saves the file as ComboBoxDemo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Path where the resulting workbook will be saved
            string outputPath = "ComboBoxDemo.xlsx";

            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // 2. Populate five items in cells A1:A5
                string[] items = { "Item One", "Item Two", "Item Three", "Item Four", "Item Five" };
                for (int i = 0; i < items.Length; i++)
                {
                    // Cells are zero‑based; row i, column 0 corresponds to A(i+1)
                    worksheet.Cells[i, 0].PutValue(items[i]);
                }

                // 3. Add a ComboBox shape to the worksheet
                // Parameters: topRow, top (pixel offset), leftColumn, left (pixel offset), height (px), width (px)
                // We'll place it starting at row 1, column 2 (i.e., cell C2) with a reasonable size.
                int topRow = 1;          // zero‑based row index
                int top = 0;             // vertical pixel offset
                int leftColumn = 2;      // zero‑based column index
                int left = 0;            // horizontal pixel offset
                int height = 30;         // height in pixels
                int width = 120;         // width of the visible combo box (in pixels)

                ComboBox comboBox = worksheet.Shapes.AddComboBox(topRow, top, leftColumn, left, height, width);

                // 4. Link the ComboBox to a cell (optional, shows selected value)
                comboBox.LinkedCell = "$B$1";

                // 5. Set the input range that supplies the items
                comboBox.InputRange = "$A$1:$A$5";

                // 6. Define the drop‑down width (list control width)
                // This controls how wide the list appears when the user clicks the arrow.
                comboBox.Width = 200; // width in pixels

                // 7. Optionally set the number of visible lines in the drop‑down
                comboBox.DropDownLines = 5;

                // 8. Save the workbook
                workbook.Save(outputPath);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("ComboBox created with five items and a custom drop‑down width.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the ComboBox:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
