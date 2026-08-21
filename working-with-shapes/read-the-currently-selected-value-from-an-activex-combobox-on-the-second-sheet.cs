// Title: C# – Read Selected Value from an ActiveX ComboBox on the Second Worksheet with Aspose.Cells
// Description: Loads an Excel workbook, accesses the second worksheet, scans its Shapes collection for a ComboBoxActiveXControl, reads the control's current Value property, and outputs the selected item. Includes handling for missing ComboBox controls.
// Keywords: Aspose.Cells | C# | .NET | ActiveX ComboBox | ComboBoxActiveXControl | read ComboBox value | second worksheet | Excel shapes collection | retrieve selected item | Excel without Interop
// Common Searches: How to get the selected item of an ActiveX ComboBox on a specific sheet using Aspose.Cells | C# code to read ComboBoxActiveXControl value from the second worksheet | Aspose.Cells iterate shapes to find ActiveX ComboBox | Read Excel ActiveX control value without opening Excel | Retrieve ComboBox selection from Excel workbook in .NET
// Developer Intent: Extract the currently selected value of an ActiveX ComboBox placed on the second worksheet of an Excel file.
// Use Cases: Capture user‑chosen options from an embedded Excel form for downstream processing. | Validate a ComboBox selection before generating a report based on workbook data. | Log the selected ComboBox entry to a database or audit trail directly from the workbook.
// AI Prompts: Generate C# code that uses Aspose.Cells to read the selected value of a ComboBoxActiveXControl on the second worksheet and gracefully handle a missing control. | Show how to loop through all shapes on a worksheet, identify multiple ComboBoxActiveXControl instances, and collect each selected value. | Explain how to programmatically change the selected value of a ComboBoxActiveXControl and save the workbook with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

// Loads an Excel workbook, accesses the second worksheet, scans its Shapes collection for a ComboBoxActiveXControl, reads the control's current Value property, and outputs the selected item. Includes handling for missing ComboBox controls.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the second worksheet (index 1)
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
            // Read the currently selected value from the ComboBox
            string selectedValue = comboBox.Value;
            Console.WriteLine("Selected value: " + selectedValue);
        }
        else
        {
            Console.WriteLine("ComboBox ActiveX control not found on the second sheet.");
        }

        // If you need to save any changes, uncomment the line below
        // workbook.Save("output.xlsx");
    }
}
