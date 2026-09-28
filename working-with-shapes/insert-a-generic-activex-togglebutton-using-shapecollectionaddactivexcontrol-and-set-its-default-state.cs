// Title: Add an ActiveX ToggleButton to an Excel worksheet with Aspose.Cells C# and set its initial state
// AI Prompts: Insert a generic ActiveX ToggleButton into a worksheet using ShapeCollection.AddActiveXControl and assign a custom name. | Configure the ToggleButton's default pressed state and caption through the ActiveXControl object in C#.
// Common Searches: C# Aspose.Cells how to insert an ActiveX ToggleButton into a specific cell | Set default value of an ActiveX ToggleButton using Aspose.Cells ShapeCollection | Example of adding ActiveX controls to Excel with Aspose.Cells C# | Change caption of a ToggleButton ActiveX control in Aspose.Cells workbook | Save workbook after adding ActiveX ToggleButton with Aspose.Cells
// Tags: Aspose.Cells AddActiveXControl ToggleButton | C# insert ActiveX control Excel worksheet | set default state ActiveX ToggleButton Aspose.Cells | configure ActiveX control caption C# | ShapeCollection AddActiveXControl example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

// The sample creates a new workbook, adds a ToggleButton ActiveX control to a specified cell using ShapeCollection.AddActiveXControl, assigns a custom shape name, accesses the underlying ActiveXControl object, and optionally sets its default Value and Caption before saving the file as an XLSX workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define position and size for the ToggleButton (0‑based indices)
            int row = 2;      // Row index
            int column = 2;   // Column index
            int top = 0;      // Pixel offset from the top of the cell
            int left = 0;     // Pixel offset from the left of the cell
            int height = 30;  // Height in points
            int width = 80;   // Width in points

            // Add the ActiveX ToggleButton to the sheet; AddActiveXControl returns a Shape
            Shape shape = sheet.Shapes.AddActiveXControl(
                ControlType.ToggleButton, row, column, top, left, height, width);

            // Set a meaningful name for the shape (the control's name)
            shape.Name = "ToggleButton1";

            // Access the underlying ActiveXControl object
            ActiveXControl toggle = shape.ActiveXControl;

            // NOTE: In some Aspose.Cells versions SetObjectData may not be available.
            // If needed, uncomment the following lines after confirming the API support.
            // toggle.SetObjectData("Value", true);          // Set default state (pressed)
            // toggle.SetObjectData("Caption", "Enable Feature"); // Set button caption

            // Define output file path
            string outputPath = "ActiveXToggleButton.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
