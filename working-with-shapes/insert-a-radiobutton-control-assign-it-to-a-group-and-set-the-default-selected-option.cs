// Title: Aspose.Cells for .NET – Add RadioButton ActiveX Controls, Group Them, and Set a Default Selection (C#)
// Description: This example creates a new Workbook, inserts two RadioButton ActiveX controls onto the first worksheet, assigns both controls the same GroupName for mutual exclusivity, sets captions, marks the first button as checked (default) and the second as unchecked, then saves the file as an .xlsx workbook.
// Keywords: Aspose.Cells | C# | RadioButton ActiveX | Excel form control | AddActiveXControl | GroupName | default selected radio button | CheckValueType | programmatic Excel UI | worksheet shapes
// Common Searches: how to add radio button ActiveX with Aspose.Cells | group radio buttons in Excel using Aspose.Cells C# | set default checked RadioButton in Aspose.Cells workbook | Aspose.Cells add ActiveX control to worksheet | C# code for Excel radio button group Aspose
// Developer Intent: Programmatically insert two RadioButton ActiveX controls into an Excel worksheet, place them in the same group, and pre‑select one as the default option.
// Use Cases: Create a survey sheet where respondents can pick a single answer from a set of options. | Design a form‑style Excel template with mutually exclusive choices for data entry. | Build a configuration workbook that opens with a predefined default setting selected.
// AI Prompts: Generate C# code with Aspose.Cells that adds three RadioButton ActiveX controls, groups them under "Choices", and sets the second button as the default selected. | Show how to change the GroupName of existing RadioButton controls and toggle their checked state using Aspose.Cells. | Explain how to read the selected RadioButton value from a saved workbook that contains ActiveX radio button groups.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

namespace AsposeCellsRadioButtonDemo
{
    // This example creates a new Workbook, inserts two RadioButton ActiveX controls onto the first worksheet, assigns both controls the same GroupName for mutual exclusivity, sets captions, marks the first button as checked (default) and the second as unchecked, then saves the file as an .xlsx workbook.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add first RadioButton ActiveX control
                Shape shape1 = worksheet.Shapes.AddActiveXControl(
                    ControlType.RadioButton, // control type
                    1,   // upper left row index
                    0,   // vertical offset (pixels) from the row
                    1,   // upper left column index
                    0,   // horizontal offset (pixels) from the column
                    100, // width (pixels)
                    30   // height (pixels)
                );
                RadioButtonActiveXControl radio1 = (RadioButtonActiveXControl)shape1.ActiveXControl;

                // Add second RadioButton ActiveX control
                Shape shape2 = worksheet.Shapes.AddActiveXControl(
                    ControlType.RadioButton,
                    1,
                    0,
                    1,
                    50,
                    100,
                    30
                );
                RadioButtonActiveXControl radio2 = (RadioButtonActiveXControl)shape2.ActiveXControl;

                // Assign both radio buttons to the same group
                radio1.GroupName = "DemoGroup";
                radio2.GroupName = "DemoGroup";

                // Set captions for clarity
                radio1.Caption = "Option A";
                radio2.Caption = "Option B";

                // Set the default selected option (first radio button)
                radio1.Value = CheckValueType.Checked;               // selected
                radio2.Value = (CheckValueType)0;                    // unchecked (fallback if enum member missing)

                // Define output file path
                string outputPath = "RadioButtonGroupDemo.xlsx";

                // Ensure the directory exists
                string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
