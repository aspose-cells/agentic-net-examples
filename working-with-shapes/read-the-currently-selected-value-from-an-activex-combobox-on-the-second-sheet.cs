// Title: Read selected value from an ActiveX ComboBox on the second worksheet using Aspose.Cells for .NET
// Description: Load an Excel workbook, access the second sheet, locate the first ActiveX ComboBox via the Shapes collection, and retrieve its currently selected item with the ComboBoxActiveXControl.Value property.
// Keywords: Aspose.Cells | ActiveX ComboBox | ComboBoxActiveXControl | C# | read selected value | second worksheet | shape collection | Excel ActiveX control | retrieve ComboBox value | Aspose.Cells .NET example
// Common Searches: Aspose.Cells read ActiveX ComboBox value on specific sheet | How to get selected item from ComboBoxActiveXControl in C# | Retrieve ActiveX ComboBox selection from second worksheet using Aspose.Cells | Iterate shapes to find ComboBoxActiveXControl Aspose.Cells | C# code to read ComboBox value from Excel sheet with Aspose
// Developer Intent: Extract the currently selected item of an ActiveX ComboBox placed on the second worksheet of an Excel file.
// Use Cases: Display or log the user's choice from an embedded form control during workbook processing. | Drive conditional logic based on the ComboBox selection while iterating rows on the second sheet. | Export the selected ComboBox value to a report or external system for auditing.
// AI Prompts: Generate C# code that reads the selected value of a ComboBoxActiveXControl on the third worksheet using Aspose.Cells. | Show how to handle multiple ActiveX ComboBoxes on a sheet and retrieve each selected value with Aspose.Cells for .NET. | Provide an example of setting a new selected item for a ComboBoxActiveXControl, saving the workbook, and verifying the change.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

// Load an Excel workbook, access the second sheet, locate the first ActiveX ComboBox via the Shapes collection, and retrieve its currently selected item with the ComboBoxActiveXControl.Value property.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the second worksheet (index 1)
        Worksheet sheet = workbook.Worksheets[1];

        // Locate the first ComboBox ActiveX control on the sheet
        ComboBoxActiveXControl comboBox = null;
        foreach (Shape shape in sheet.Shapes)
        {
            if (shape.ActiveXControl is ComboBoxActiveXControl)
            {
                comboBox = (ComboBoxActiveXControl)shape.ActiveXControl;
                break; // Assuming only one ComboBox is needed
            }
        }

        if (comboBox != null)
        {
            // Read the currently selected value of the ComboBox
            string selectedValue = comboBox.Value; // For ActiveX ComboBox, Value holds the selected item
            Console.WriteLine("Selected value from ComboBox: " + selectedValue);
        }
        else
        {
            Console.WriteLine("No ComboBox ActiveX control found on the second sheet.");
        }

        // If you need to save any changes, uncomment the line below
        // workbook.Save("output.xlsx");
    }
}
