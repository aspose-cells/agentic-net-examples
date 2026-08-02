// Title: Add grouped ActiveX RadioButton controls and set a default selection with Aspose.Cells for .NET (C#)
// Description: This example creates a new Workbook, inserts two RadioButton ActiveX controls into the first worksheet, assigns them the same GroupName for mutual exclusivity, sets custom captions, marks the first button as checked (default), and saves the file as an .xlsx workbook.
// Keywords: Aspose.Cells | C# | .NET | ActiveX RadioButton | RadioButtonActiveXControl | GroupName | default checked | Excel form control | add shape to worksheet | multiple choice Excel form | Excel UI automation
// Common Searches: Aspose.Cells add RadioButton ActiveX control C# | group RadioButtonActiveXControl in Excel with Aspose.Cells | set default checked RadioButton in Aspose.Cells workbook | how to assign GroupName to RadioButton in Aspose.Cells | create mutually exclusive radio buttons using Aspose.Cells .NET
// Developer Intent: Insert two RadioButton ActiveX controls, place them in the same group, and make the first button the default selected option.
// Use Cases: Design a single‑choice questionnaire inside an Excel sheet. | Build a settings panel with Yes/No options that require exclusive selection. | Create a simple Excel form where users pick one preference before proceeding.
// AI Prompts: Generate C# code that adds three RadioButton ActiveX controls to a worksheet with Aspose.Cells, groups them, and sets the second button as the default checked. | Show how to change the GroupName of existing RadioButtonActiveXControl objects in an Aspose.Cells workbook to form separate exclusive groups. | Provide an example that reads the selected RadioButton value from an Aspose.Cells workbook after it is opened in Excel.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

namespace AsposeCellsRadioButtonDemo
{
    // This example creates a new Workbook, inserts two RadioButton ActiveX controls into the first worksheet, assigns them the same GroupName for mutual exclusivity, sets custom captions, marks the first button as checked (default), and saves the file as an .xlsx workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add the first RadioButton ActiveX control
                Shape shape1 = worksheet.Shapes.AddActiveXControl(
                    ControlType.RadioButton, // Control type
                    1,   // Upper left row index
                    0,   // Row offset (pixels)
                    1,   // Upper left column index
                    0,   // Column offset (pixels)
                    100, // Width (pixels)
                    30   // Height (pixels)
                );
                RadioButtonActiveXControl radio1 = (RadioButtonActiveXControl)shape1.ActiveXControl;

                // Add the second RadioButton ActiveX control
                Shape shape2 = worksheet.Shapes.AddActiveXControl(
                    ControlType.RadioButton,
                    2,   // Upper left row index (different row)
                    0,
                    1,
                    0,
                    100,
                    30
                );
                RadioButtonActiveXControl radio2 = (RadioButtonActiveXControl)shape2.ActiveXControl;

                // Assign both radio buttons to the same group so they are mutually exclusive
                const string groupName = "OptionsGroup";
                radio1.GroupName = groupName;
                radio2.GroupName = groupName;

                // Set captions for the options
                radio1.Caption = "Option A";
                radio2.Caption = "Option B";

                // Set the default selected option (first radio button)
                radio1.Value = CheckValueType.Checked;               // Selected
                // Use explicit cast for unchecked value to avoid enum version issues
                radio2.Value = (CheckValueType)0;                     // Not selected

                // Define output file path
                string outputPath = "RadioButtonGroupDemo.xlsx";

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
