// Title: Cast Shape.ActiveXControl to CheckBoxActiveXControl in Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds a CheckBox ActiveX control via Worksheet.Shapes.AddActiveXControl, verifies the shape contains an ActiveXControl, casts it to CheckBoxActiveXControl, reads the Font.Name, updates the Caption, and saves the file.
// Keywords: Aspose.Cells ActiveXControl cast | CheckBoxActiveXControl C# | Shape.ActiveXControl example | modify ActiveX control properties Aspose.Cells | .NET spreadsheet ActiveX control
// Common Searches: How to cast Shape.ActiveXControl to a specific type in Aspose.Cells | Change CheckBox caption using Aspose.Cells C# | Read font name of an ActiveX CheckBox in a worksheet | Access ActiveX controls added to Excel with Aspose.Cells
// Developer Intent: Retrieve a Shape's ActiveXControl, safely cast it to its concrete CheckBoxActiveXControl class, and manipulate its properties.
// Use Cases: Extract and display the font name of a CheckBox ActiveX control placed on a sheet | Programmatically change the caption of a CheckBox after insertion | Validate the presence of an ActiveX control before casting to prevent runtime errors
// AI Prompts: Write C# code that adds a ComboBox ActiveX control to a worksheet, accesses it through Shape.ActiveXControl, casts to ComboBoxActiveXControl, and populates its items. | Explain best practices for safely casting Shape.ActiveXControl to a specific control type and handling invalid casts in Aspose.Cells. | Provide a step‑by‑step guide to iterate over all worksheet shapes, detect ActiveX controls, and modify a property based on the control's concrete type.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a CheckBox ActiveX control via Worksheet.Shapes.AddActiveXControl, verifies the shape contains an ActiveXControl, casts it to CheckBoxActiveXControl, reads the Font.Name, updates the Caption, and saves the file.
    public class AccessActiveXControlDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and obtain the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a CheckBox ActiveX control to the worksheet
                Shape shape = worksheet.Shapes.AddActiveXControl(
                    ControlType.CheckBox, // type of control
                    1,    // top row index
                    0,    // vertical offset in pixels
                    1,    // left column index
                    0,    // horizontal offset in pixels
                    100,  // width in pixels
                    30    // height in pixels
                );

                // Verify that the shape contains an ActiveX control
                if (shape.ActiveXControl != null)
                {
                    // Cast the generic ActiveXControl to its specific CheckBox type
                    CheckBoxActiveXControl checkBox = (CheckBoxActiveXControl)shape.ActiveXControl;

                    // Access a property specific to CheckBoxActiveXControl (e.g., Font name)
                    string fontName = checkBox.Font.Name;
                    Console.WriteLine($"CheckBox font name: {fontName}");

                    // Modify a property specific to the CheckBox control
                    checkBox.Caption = "I Agree";
                }

                // Save the workbook to persist the added control
                workbook.Save("AccessActiveXControlDemo.xlsx");
                Console.WriteLine("Workbook saved as AccessActiveXControlDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            AccessActiveXControlDemo.Run();
        }
    }
}
