using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsLinkedShapeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the ListBox input range
            for (int i = 0; i < 6; i++)
            {
                worksheet.Cells[i, 0].PutValue(i + 1); // A1:A6 = 1..6
            }

            // Add a ListBox shape to the worksheet
            Shape listBoxShape = worksheet.Shapes.AddListBox(2, 0, 2, 0, 130, 130);

            // Define the input range (items) and the linked cell (selected value)
            listBoxShape.SetInputRange("$A$1:$A$6", false, false);
            listBoxShape.SetLinkedCell("$A$12", false, true);

            // Initial selection: set linked cell to 3 and refresh the shape
            worksheet.Cells["A12"].PutValue(3);
            listBoxShape.UpdateSelectedValue();

            // Verify the selection (optional)
            ListBox listBox = (ListBox)listBoxShape;
            if (listBox.IsSelected(2)) // zero‑based index, 2 corresponds to value 3
            {
                Console.WriteLine("Option 3 is selected after first update.");
            }

            // Change the linked cell value to 4 and refresh the shape again
            worksheet.Cells["A12"].PutValue(4);
            listBoxShape.UpdateSelectedValue();

            // Verify the new selection
            if (listBox.IsSelected(3)) // index 3 corresponds to value 4
            {
                Console.WriteLine("Option 4 is selected after second update.");
            }

            // Save the workbook to a file
            workbook.Save("LinkedShapeDemo.xlsx");
        }
    }
}