using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

class RadioButtonGroupExample
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
                ControlType.RadioButton, // control type
                1,   // top row index
                0,   // vertical offset (pixels)
                1,   // left column index
                0,   // horizontal offset (pixels)
                100, // height (pixels)
                30   // width (pixels)
            );
            RadioButtonActiveXControl radio1 = (RadioButtonActiveXControl)shape1.ActiveXControl;
            radio1.GroupName = "MyRadioGroup";          // assign to a group
            radio1.Caption = "Option A";                // set display text
            radio1.Value = CheckValueType.Checked;      // default selected option

            // Add the second RadioButton ActiveX control
            Shape shape2 = worksheet.Shapes.AddActiveXControl(
                ControlType.RadioButton,
                1,
                0,
                1,
                50,   // place it a bit lower
                100,
                30
            );
            RadioButtonActiveXControl radio2 = (RadioButtonActiveXControl)shape2.ActiveXControl;
            radio2.GroupName = "MyRadioGroup";          // same group makes them mutually exclusive
            radio2.Caption = "Option B";

            // The enum member 'Unchecked' may not be available in some versions;
            // casting the underlying integer (0) to the enum achieves the same effect.
            radio2.Value = (CheckValueType)0;           // not selected by default

            // Define output file path
            string outputPath = "RadioButtonGroupDemo.xlsx";

            // Save the workbook with the radio buttons
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}