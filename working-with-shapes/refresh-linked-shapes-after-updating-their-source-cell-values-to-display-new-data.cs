using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsLinkedShapeRefresh
{
    public class RefreshLinkedShapesDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a ListBox shape (dropdown) to the worksheet
            // Parameters: upper left row, upper left column, top, left, width, height
            Shape listBoxShape = worksheet.Shapes.AddListBox(2, 0, 2, 0, 130, 130);

            // Define the range that provides the list items
            listBoxShape.SetInputRange("$A$1:$A$5", false, false);

            // Link the selected value of the ListBox to a cell (e.g., A10)
            listBoxShape.SetLinkedCell("$A$10", false, true);

            // Populate the input range with sample items
            worksheet.Cells["A1"].Value = "Option 1";
            worksheet.Cells["A2"].Value = "Option 2";
            worksheet.Cells["A3"].Value = "Option 3";
            worksheet.Cells["A4"].Value = "Option 4";
            worksheet.Cells["A5"].Value = "Option 5";

            // Set an initial value in the linked cell (select "Option 2")
            worksheet.Cells["A10"].Value = "Option 2";

            // Refresh the shape so it reflects the linked cell value
            listBoxShape.UpdateSelectedValue();

            // Change the linked cell value to a different option (e.g., "Option 4")
            worksheet.Cells["A10"].Value = "Option 4";

            // Refresh the shape again to display the new selection
            listBoxShape.UpdateSelectedValue();

            // Save the workbook to verify the result
            string outputPath = "LinkedShapeRefreshDemo.xlsx";

            // Ensure we can write to the output location
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
    }
}