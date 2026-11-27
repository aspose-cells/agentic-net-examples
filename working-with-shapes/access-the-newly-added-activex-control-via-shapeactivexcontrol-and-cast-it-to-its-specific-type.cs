// Title: Retrieve a Shape's ActiveXControl and cast it to CheckBoxActiveXControl with Aspose.Cells for .NET
// Common Searches: how to get the ActiveXControl object from a Shape in Aspose.Cells C# | casting a generic ActiveXControl to CheckBoxActiveXControl in .NET | setting caption, linked cell and checked state for an ActiveX checkbox with Aspose.Cells | enabling and making visible an ActiveX checkbox added to an Excel worksheet using Aspose.Cells
// Tags: retrieve ActiveXControl from Shape Aspose.Cells | cast generic ActiveXControl to specific CheckBoxActiveXControl | configure CheckBoxActiveXControl caption and linked cell | set CheckBoxActiveXControl checked value Aspose.Cells | enable and show ActiveX checkbox in worksheet Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

// The example creates a workbook, adds a CheckBox ActiveX control via Shape.AddActiveXControl, verifies the control exists, casts the generic ActiveXControl to CheckBoxActiveXControl, configures its caption, linked cell, checked state, font, and visibility, then saves the file as an XLSX workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a CheckBox ActiveX control to the worksheet.
            // Parameters: ControlType, upper left row, top offset (pixels),
            // upper left column, left offset (pixels), width (pixels), height (pixels)
            Shape shape = worksheet.Shapes.AddActiveXControl(
                ControlType.CheckBox,
                1, 0,   // upper left row, top offset
                1, 0,   // upper left column, left offset
                120, 30 // width, height
            );

            // Verify that the shape actually contains an ActiveX control
            if (shape.ActiveXControl == null)
            {
                Console.WriteLine("The shape does not contain an ActiveX control.");
                return;
            }

            // Cast the generic ActiveXControl to the specific CheckBoxActiveXControl type
            CheckBoxActiveXControl checkBox = shape.ActiveXControl as CheckBoxActiveXControl;
            if (checkBox == null)
            {
                Console.WriteLine("Failed to cast the ActiveXControl to CheckBoxActiveXControl.");
                return;
            }

            // Set properties specific to CheckBoxActiveXControl
            checkBox.Caption = "I agree to the terms";
            checkBox.LinkedCell = "B2";

            // Set the initial checked state using the correct enum type
            checkBox.Value = CheckValueType.Checked;

            // Change the font of the caption
            checkBox.Font.Name = "Calibri";
            checkBox.Font.Size = 11;

            // Enable the control and make it visible
            checkBox.IsEnabled = true;
            checkBox.IsVisible = true;

            // Output some information to the console
            Console.WriteLine("CheckBox ActiveX control added successfully.");
            Console.WriteLine($"Caption: {checkBox.Caption}");
            Console.WriteLine($"Linked Cell: {checkBox.LinkedCell}");
            Console.WriteLine($"Font: {checkBox.Font.Name}, Size: {checkBox.Font.Size}");
            Console.WriteLine($"Checked Value: {checkBox.Value}");

            // Save the workbook to an XLSX file
            string outputPath = "ActiveXCheckBoxDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
