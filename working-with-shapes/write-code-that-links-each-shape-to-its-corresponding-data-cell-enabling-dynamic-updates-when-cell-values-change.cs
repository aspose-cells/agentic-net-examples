using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class ShapeLinkingDemo
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
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data cells that will be linked to shapes
            for (int row = 0; row < 5; row++)
            {
                sheet.Cells[row, 0].Value = row + 1;               // Column A values (1..5)
                sheet.Cells[row, 1].Value = $"Item {row + 1}";    // Column B values (text)
            }

            // Add a ListBox shape and link it to cell A1 (value determines selected index)
            Shape listBoxShape = sheet.Shapes.AddListBox(2, 0, 2, 0, 120, 100);
            ((ListBox)listBoxShape).SetInputRange("$B$1:$B$5", false, false);
            listBoxShape.SetLinkedCell("$A$1", false, true); // Linked cell A1

            // Add a CheckBox shape and link it to cell A2 (TRUE/FALSE)
            Shape checkBoxShape = sheet.Shapes.AddCheckBox(5, 0, 5, 0, 100, 30);
            checkBoxShape.SetLinkedCell("$A$2", false, true); // Linked cell A2

            // Add a Spinner shape and link it to cell A3 (numeric value)
            Shape spinnerShape = sheet.Shapes.AddSpinner(8, 0, 8, 0, 100, 30);
            spinnerShape.SetLinkedCell("$A$3", false, true); // Linked cell A3

            // Initial update: shapes read the current linked cell values
            listBoxShape.UpdateSelectedValue();
            checkBoxShape.UpdateSelectedValue();
            spinnerShape.UpdateSelectedValue();

            // Change linked cell values to demonstrate dynamic updates
            sheet.Cells["A1"].Value = 3;      // Select third item in ListBox
            sheet.Cells["A2"].Value = true;  // Check the CheckBox
            sheet.Cells["A3"].Value = 7;      // Set Spinner value

            // Update shapes again so they reflect the new cell values
            listBoxShape.UpdateSelectedValue();
            checkBoxShape.UpdateSelectedValue();
            spinnerShape.UpdateSelectedValue();

            // Save the workbook
            string outputPath = "ShapeLinkingDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}