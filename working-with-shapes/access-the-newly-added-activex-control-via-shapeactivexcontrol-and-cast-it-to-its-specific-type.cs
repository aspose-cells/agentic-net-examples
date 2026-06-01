using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

namespace AsposeCellsExamples
{
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

                // Verify that the control was added
                if (shape.ActiveXControl != null)
                {
                    // Cast the generic ActiveXControl to its specific type
                    CheckBoxActiveXControl checkBox = (CheckBoxActiveXControl)shape.ActiveXControl;

                    // Access a property specific to CheckBoxActiveXControl
                    string fontName = checkBox.Font.Name;
                    Console.WriteLine($"CheckBox Font Name: {fontName}");

                    // Modify some properties of the CheckBox
                    checkBox.Caption = "Accept Terms";
                    checkBox.IsEnabled = true;
                }

                // Save the workbook to a file
                string outputPath = "AccessActiveXControlDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            AccessActiveXControlDemo.Run();
        }
    }
}