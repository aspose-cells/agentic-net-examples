using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a ListBox shape (a form control) to the worksheet
            Shape listBox = worksheet.Shapes.AddListBox(2, 0, 2, 0, 130, 130);

            // Initially link the shape to cell A1
            listBox.SetLinkedCell("A1", false, true);
            Console.WriteLine("Initial LinkedCell: " + listBox.LinkedCell); // Should output A1

            // Change the linked cell from A1 to B2
            listBox.SetLinkedCell("B2", false, true);
            Console.WriteLine("Changed LinkedCell: " + listBox.LinkedCell); // Should output B2

            // Put test values into the cells
            worksheet.Cells["A1"].Value = "Value in A1";
            worksheet.Cells["B2"].Value = "Value in B2";

            // Update the shape's selected value based on the new linked cell (B2)
            listBox.UpdateSelectedValue();

            // Verify that the shape reflects the value from the linked cell
            string linkedCellAddress = listBox.LinkedCell;               // e.g., "B2"
            // Ensure the address is in a format accepted by Cells[]
            string cleanAddress = linkedCellAddress.Replace("$", string.Empty);
            object linkedCellValue = worksheet.Cells[cleanAddress].Value;
            Console.WriteLine($"Linked cell ({linkedCellAddress}) value: {linkedCellValue}");

            // Save the workbook (optional, demonstrates full lifecycle)
            string outputPath = "LinkedCellDemo.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Runtime error: {ex.Message}");
        }
    }
}