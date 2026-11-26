// Title: How to set the position and size of an ActiveX CheckBox control with Shape properties in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# move ActiveX CheckBox to a specific row and column | Change pixel offset of an ActiveX control shape in an Excel worksheet using Aspose.Cells | Resize an ActiveX control after adding it with Shapes.AddActiveXControl in .NET | Set UpperLeftRow and UpperLeftColumn for an ActiveX shape via Aspose.Cells API | Adjust Top, Left, Width, and Height of an ActiveX control in a generated XLSX file | Programmatically assign a linked cell to an ActiveX CheckBox using Aspose.Cells C#
// Tags: shape upperleftrow property Aspose.Cells | shape upperleftcolumn property Aspose.Cells | shape top left pixel offset Aspose.Cells | shape width height adjustment Aspose.Cells | addactivexcontrol resize shape C# | activex checkbox linkedcell Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

namespace ActiveXControlPositionSizeDemo
{
    // The example creates a new workbook, inserts an ActiveX CheckBox via Worksheet.Shapes.AddActiveXControl, then uses the Shape object's UpperLeftRow, UpperLeftColumn, Top, Left, Width, and Height properties to reposition and resize the control. It also sets the CheckBox caption, checked state, and linked cell before saving the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // -----------------------------------------------------------------
                // 1. Insert an ActiveX CheckBox control at an initial location.
                // -----------------------------------------------------------------
                Shape shape = worksheet.Shapes.AddActiveXControl(
                    ControlType.CheckBox, // type of control
                    1,                    // upper left row index (0‑based)
                    0,                    // vertical offset in pixels
                    1,                    // upper left column index (0‑based)
                    0,                    // horizontal offset in pixels
                    100,                  // initial width in pixels
                    30);                  // initial height in pixels

                // -----------------------------------------------------------------
                // 2. Adjust the position and size of the inserted shape.
                // -----------------------------------------------------------------
                shape.UpperLeftRow = 5;   // row where the shape's top edge starts
                shape.UpperLeftColumn = 3; // column where the shape's left edge starts
                shape.Top = 12;           // 12 pixels down from the top of the cell
                shape.Left = 8;           // 8 pixels right from the left of the cell
                shape.Width = 150;        // new width (pixels)
                shape.Height = 45;        // new height (pixels)

                // -----------------------------------------------------------------
                // 3. Set properties of the underlying ActiveX CheckBox control.
                // -----------------------------------------------------------------
                if (shape.ActiveXControl is CheckBoxActiveXControl checkBox)
                {
                    checkBox.Caption = "I agree";
                    // Use the enum CheckValueType to set the checked state
                    checkBox.Value = CheckValueType.Checked;
                    checkBox.LinkedCell = "B2";
                }

                // -----------------------------------------------------------------
                // 4. Save the workbook.
                // -----------------------------------------------------------------
                string outputPath = "ActiveXControlPositionSize.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("ActiveX control position and size have been set:");
                Console.WriteLine($"  UpperLeftRow = {shape.UpperLeftRow}");
                Console.WriteLine($"  UpperLeftColumn = {shape.UpperLeftColumn}");
                Console.WriteLine($"  Top (px) = {shape.Top}");
                Console.WriteLine($"  Left (px) = {shape.Left}");
                Console.WriteLine($"  Width (px) = {shape.Width}");
                Console.WriteLine($"  Height (px) = {shape.Height}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
