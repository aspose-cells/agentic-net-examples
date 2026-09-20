// Title: Create a ComboBox shape in an Excel worksheet, add five items, and set its drop‑down width using Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to insert a ComboBox shape into the first worksheet, add five list items, set the drop‑down width to 150 points, and configure the DropDownLines property to show all items. | Generate a complete Aspose.Cells .NET example that creates a workbook, places a ComboBox at a specific cell range, populates it with custom entries, adjusts the dropdown width, and saves the file as an .xlsx document.
// Common Searches: c# aspocells add combobox shape to worksheet | aspocells set combobox dropdown width and visible lines | how to populate combobox list items using aspocells .net | example creating combobox with five items in excel using aspocells | aspocells combo box DropDownLines property usage
// Tags: Aspose.Cells add ComboBox shape | Aspose.Cells populate ComboBox items | Aspose.Cells set ComboBox drop-down width | Aspose.Cells configure ComboBox DropDownLines | Aspose.Cells save workbook as Xlsx

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, adds a ComboBox shape to the first worksheet within a defined cell range, populates it with five items, sets the DropDownLines property to display five rows, specifies a 150‑point drop‑down width, and saves the workbook as ComboBoxExample.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the position of the ComboBox (zero‑based indices)
            int upperLeftRow = 2;      // 3rd row
            int upperLeftColumn = 1;   // 2nd column
            int lowerRightRow = 4;     // 5th row
            int lowerRightColumn = 3;  // 4th column

            // Add a ComboBox shape. Height and width are required parameters (in points).
            ComboBox comboBox = sheet.Shapes.AddComboBox(
                upperLeftRow, upperLeftColumn,
                lowerRightRow, lowerRightColumn,
                30,   // height
                150); // width

            // Populate the ComboBox with items.
            // Using dynamic to accommodate possible API variations across Aspose.Cells versions.
            dynamic cb = comboBox;
            cb.ListItems.Add("Item 1");
            cb.ListItems.Add("Item 2");
            cb.ListItems.Add("Item 3");
            cb.ListItems.Add("Item 4");
            cb.ListItems.Add("Item 5");

            // Optionally set the number of visible items in the drop‑down list
            comboBox.DropDownLines = 5;

            // Save the workbook
            string outputPath = "ComboBoxExample.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
