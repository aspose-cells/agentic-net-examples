using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class FilterErrorLinkedShapes
{
    static void Main()
    {
        // Load the workbook with options to ignore useless shapes
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.IgnoreUselessShapes = true;
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            ShapeCollection shapes = sheet.Shapes;

            // Examine each shape in the worksheet
            for (int i = 0; i < shapes.Count; i++)
            {
                Shape shape = shapes[i];

                // Retrieve the linked cell address (if the shape is linked)
                string linkedCellAddress = shape.GetLinkedCell(true, true);

                // Proceed only if the shape has a linked cell
                if (!string.IsNullOrEmpty(linkedCellAddress))
                {
                    // Access the cell using the address
                    Cell linkedCell = sheet.Cells[linkedCellAddress];

                    // Check whether the linked cell contains an error value
                    if (linkedCell != null && linkedCell.IsErrorValue)
                    {
                        // Log details of the problematic shape
                        Console.WriteLine($"Worksheet: {sheet.Name}, Shape Index: {i}, Shape Name: {shape.Name}, Linked Cell: {linkedCellAddress}, Error Value: {linkedCell.Value}");
                    }
                }
            }
        }

        // Save the workbook (unchanged) after processing
        workbook.Save("output.xlsx");
    }
}