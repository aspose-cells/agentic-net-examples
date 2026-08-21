// Title: Create an ActiveX ComboBox in an Excel sheet, bind it to cells, and set its drop‑down width with Aspose.Cells for .NET (C#)
// Description: This example shows how to generate a workbook, write five items into A1:A5, insert a ComboBox ActiveX control, link the control to the range via ListFillRange, configure the drop‑down list width (ListWidth) and visible rows (ListRows), and save the file as ComboBoxWithDropDownWidth.xlsx using Aspose.Cells for C#.
// Keywords: Aspose.Cells ComboBox ActiveX | C# set ComboBox ListWidth | populate ComboBox from cell range Aspose | Excel ActiveX control dropdown width | ListFillRange Aspose.Cells | ComboBox ListRows property
// Common Searches: how to add a ComboBox ActiveX control with Aspose.Cells | set dropdown width for ComboBox in Aspose.Cells .NET | bind ComboBox list to Excel range using Aspose | adjust visible rows of ComboBox list in Aspose.Cells
// Developer Intent: Insert a ComboBox ActiveX control, fill it from worksheet cells, and customize its drop‑down dimensions programmatically.
// Use Cases: Design a form‑like worksheet where users pick options from a wider ComboBox list for better readability. | Generate a template that lists categories from a range in a ComboBox with a predefined number of visible rows. | Automate report creation that includes a ComboBox with a custom drop‑down width to improve the UI experience.
// AI Prompts: Show C# code to add an ActiveX ComboBox to an Excel sheet and bind it to a cell range using Aspose.Cells. | Provide an example that changes the ListWidth and ListRows of a ComboBoxActiveXControl after insertion. | Explain how to configure ListFillRange, ListWidth, and ListRows for a ComboBoxActiveXControl in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

// This example shows how to generate a workbook, write five items into A1:A5, insert a ComboBox ActiveX control, link the control to the range via ListFillRange, configure the drop‑down list width (ListWidth) and visible rows (ListRows), and save the file as ComboBoxWithDropDownWidth.xlsx using Aspose.Cells for C#.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate items for the combo box
            sheet.Cells["A1"].PutValue("Item 1");
            sheet.Cells["A2"].PutValue("Item 2");
            sheet.Cells["A3"].PutValue("Item 3");
            sheet.Cells["A4"].PutValue("Item 4");
            sheet.Cells["A5"].PutValue("Item 5");

            // Add a ComboBox ActiveX control
            // Parameters: control type, upper left row, upper left column, top offset, left offset, width, height
            Shape shape = sheet.Shapes.AddActiveXControl(
                ControlType.ComboBox,
                0,   // upper left row (0‑based)
                0,   // upper left column (0‑based)
                5,   // top offset in pixels
                5,   // left offset in pixels
                100, // width in pixels
                30   // height in pixels
            );

            // Retrieve the ComboBox control object
            ComboBoxActiveXControl comboBox = (ComboBoxActiveXControl)shape.ActiveXControl;

            // Link the combo box to the range containing the items
            comboBox.ListFillRange = "A1:A5";

            // Set dropdown list width (points) and visible rows
            comboBox.ListWidth = 150;
            comboBox.ListRows = 5;

            // Save the workbook
            string outputPath = "ComboBoxWithDropDownWidth.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
