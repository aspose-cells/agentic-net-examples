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

            // Populate the list values in cells A1:A3
            worksheet.Cells["A1"].Value = "Option1";
            worksheet.Cells["A2"].Value = "Option2";
            worksheet.Cells["A3"].Value = "Option3";

            // Create a data validation on cell B1 that shows an in‑cell dropdown list
            Validation validation = worksheet.Cells["B1"].GetValidation();
            validation.Type = ValidationType.List;
            validation.Formula1 = "$A$1:$A$3";   // reference the list range
            validation.InCellDropDown = true;   // enable the dropdown

            // Add a ListBox shape to the worksheet
            Shape listBoxShape = worksheet.Shapes.AddListBox(2, 0, 2, 0, 130, 130);
            if (listBoxShape == null)
                throw new InvalidOperationException("Failed to create ListBox shape.");

            // Set the input range of the ListBox to the same list range
            listBoxShape.SetInputRange("$A$1:$A$3", false, false);

            // Link the ListBox to cell B1 so its selected value reflects the cell value
            listBoxShape.SetLinkedCell("$B$1", false, true);

            // Set an initial value in the linked cell (e.g., "Option2")
            worksheet.Cells["B1"].Value = "Option2";

            // Update the shape's selected item based on the linked cell value
            listBoxShape.UpdateSelectedValue();

            // Save the workbook
            string outputPath = "ShapeLinkedValidationDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}