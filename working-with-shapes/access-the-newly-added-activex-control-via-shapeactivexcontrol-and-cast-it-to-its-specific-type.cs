// Title: Retrieve an ActiveX button from a worksheet Shape and cast it to its specific type with Aspose.Cells for .NET
// Common Searches: Aspose.Cells get ActiveX control from Shape and cast to Button in C# | C# retrieve ActiveX button object using Shape.ActiveXControl property Aspose.Cells | How to access the underlying ActiveX control of a worksheet button with Aspose.Cells | Casting Shape.ActiveXControl to a concrete control type in Aspose.Cells .NET | Read properties of an ActiveX button added to Excel using Aspose.Cells
// Tags: Shape.ActiveXControl retrieval Aspose.Cells | cast ActiveX control to concrete type C# | access worksheet button ActiveX object .NET | Aspose.Cells ActiveX button manipulation | working with ActiveX shapes in Excel using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds an ActiveX button shape to a worksheet, then shows how to obtain the button's underlying ActiveX control via the Shape.ActiveXControl property and cast it to its specific type for further processing.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a button shape at cell B2 (row 1, column 1)
            // Parameters: upper left row, column, top offset, left offset, width, height (all in points)
            Shape buttonShape = worksheet.Shapes.AddButton(
                1,    // upper left row (0‑based)
                1,    // upper left column (0‑based)
                0,    // top offset in points
                0,    // left offset in points
                100,  // width in points
                30);  // height in points

            // Set the button caption
            buttonShape.Text = "Click Me";

            // Optionally, enable the button (form controls are enabled by default)
            // No explicit property needed; kept for clarity
            // buttonShape.Enabled = true; // not required

            // Define output file path
            string outputPath = "ActiveXButton.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
