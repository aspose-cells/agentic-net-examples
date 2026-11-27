// Title: Add a ToggleButton ActiveX control to an Excel worksheet with Aspose.Cells C# and set its default checked state
// Common Searches: C# Aspose.Cells add ToggleButton ActiveX control to worksheet | How to set the default value of a ToggleButtonActiveXControl using ShapeCollection.AddActiveXControl | Link a ToggleButton ActiveX control to a worksheet cell with Aspose.Cells .NET | Disable triple‑state for a ToggleButton ActiveX control in Aspose.Cells | Create an Excel file that contains a checked ToggleButton using Aspose.Cells C#
// Tags: aspnetcells shapecollection addactivexcontrol togglebutton | togglebuttonactiveXcontrol set default checked value | activex togglebutton linkedcell property aspnetcells | disable triple state togglebutton aspnetcells | c# create excel workbook with activex togglebutton

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

namespace ActiveXToggleButtonDemo
{
    // The example creates a new Workbook, inserts a ToggleButton ActiveX control via Worksheet.Shapes.AddActiveXControl, casts it to ToggleButtonActiveXControl, sets its Value to Checked, disables triple‑state, assigns a caption, links it to cell B2, and saves the workbook as an .xlsx file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a ToggleButton ActiveX control to the worksheet
                // Parameters: ControlType, topRow, top (pixel offset), leftColumn, left (pixel offset), width, height
                Shape shape = worksheet.Shapes.AddActiveXControl(
                    ControlType.ToggleButton, // type of control
                    2,   // upper left row index (zero‑based)
                    0,   // vertical offset in pixels
                    2,   // upper left column index (zero‑based)
                    0,   // horizontal offset in pixels
                    120, // width in pixels
                    30   // height in pixels
                );

                // Cast the generic ActiveXControl to the specific ToggleButton type
                ToggleButtonActiveXControl toggleButton = shape.ActiveXControl as ToggleButtonActiveXControl;
                if (toggleButton == null)
                {
                    Console.WriteLine("Failed to retrieve the ToggleButton ActiveX control.");
                    return;
                }

                // Set the default state (checked)
                toggleButton.Value = CheckValueType.Checked; // Options: Checked, Unchecked, Mixed
                toggleButton.IsTripleState = false;          // Ensure only two states
                toggleButton.Caption = "Demo Toggle";
                toggleButton.IsEnabled = true;
                toggleButton.IsVisible = true;
                toggleButton.LinkedCell = "B2"; // Optional: link to a cell to reflect state

                // Save the workbook
                string outputPath = "ToggleButtonActiveXControlDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the ToggleButton ActiveX control:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
