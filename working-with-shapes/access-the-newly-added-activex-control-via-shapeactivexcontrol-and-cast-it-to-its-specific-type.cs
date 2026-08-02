// Title: How to Retrieve and Cast an ActiveX CheckBox from a Shape using Aspose.Cells for .NET
// Description: This example creates a workbook, inserts a CheckBox ActiveX control via Worksheet.Shapes.AddActiveXControl, accesses the generic control through Shape.ActiveXControl, safely casts it to CheckBoxActiveXControl, reads its Font.Name, updates the Caption, and saves the file.
// Keywords: Aspose.Cells ActiveXControl | Shape.ActiveXControl cast | CheckBoxActiveXControl .NET | add ActiveX checkbox Aspose | modify ActiveX properties Aspose.Cells
// Common Searches: Aspose.Cells get specific ActiveX control from shape | cast Shape.ActiveXControl to CheckBoxActiveXControl C# | change caption of ActiveX checkbox Aspose.Cells | read font name of ActiveX checkbox shape
// Developer Intent: Obtain the concrete ActiveX control attached to a Shape and manipulate its type‑specific properties in a .NET workbook.
// Use Cases: Insert a CheckBox ActiveX control and later change its caption programmatically. | Read the font name of a checkbox after casting the generic control. | Validate Shape.ActiveXControl for null before casting to avoid exceptions.
// AI Prompts: Generate C# code that accesses Shape.ActiveXControl, checks for null, and casts it to the appropriate ActiveX type in Aspose.Cells. | Show how to loop through all worksheet shapes, detect ActiveX controls, and update properties based on each control’s specific class.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

// This example creates a workbook, inserts a CheckBox ActiveX control via Worksheet.Shapes.AddActiveXControl, accesses the generic control through Shape.ActiveXControl, safely casts it to CheckBoxActiveXControl, reads its Font.Name, updates the Caption, and saves the file.
public class AccessActiveXControlDemo
{
    public static void Run()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a CheckBox ActiveX control to the worksheet
            Shape shape = worksheet.Shapes.AddActiveXControl(
                ControlType.CheckBox, // type of control
                1,    // topRow index
                0,    // vertical offset in pixels
                1,    // leftColumn index
                0,    // horizontal offset in pixels
                100,  // width in pixels
                30    // height in pixels
            );

            // Access the ActiveX control via the Shape.ActiveXControl property
            if (shape.ActiveXControl != null)
            {
                // Cast the generic ActiveXControl to its specific type (CheckBoxActiveXControl)
                CheckBoxActiveXControl checkBox = (CheckBoxActiveXControl)shape.ActiveXControl;

                // Use a property specific to CheckBoxActiveXControl
                string fontName = checkBox.Font.Name;
                Console.WriteLine($"CheckBox Font: {fontName}");

                // Example: set the caption of the checkbox
                checkBox.Caption = "Accept Terms";
            }

            // Save the workbook to a file
            string outputPath = "ActiveXControlAccessDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
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
