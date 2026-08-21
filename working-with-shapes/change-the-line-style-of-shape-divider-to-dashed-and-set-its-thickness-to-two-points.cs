// Title: Aspose.Cells for .NET – Apply Dashed 2‑Point Line Style to a Shape Called "Divider"
// Description: The sample creates a workbook, inserts a line shape named "Divider", retrieves it by its identifier, changes the line's dash pattern to a dash style and sets the weight to 2 points, then writes the result to Output.xlsx using the Aspose.Cells C# API.
// Keywords: Aspose.Cells line shape dash style | C# set line weight Aspose.Cells | modify shape properties .NET | add line shape Aspose.Cells | MsoLineDashStyle enumeration | shape line thickness example | format line shape Aspose.Cells | Aspose.Cells shape formatting C# | change line dash pattern Aspose.Cells
// Common Searches: Aspose.Cells change line dash style C# | Set line weight to 2 points for a shape in Aspose.Cells | Retrieve a shape by name Aspose.Cells .NET | How to format a line shape using Aspose.Cells | C# Aspose.Cells line shape properties example
// Developer Intent: Configure the "Divider" line shape to use a dashed pattern and a 2‑point thickness.
// Use Cases: Add a visual separator in an automatically generated report and style it with a dashed 2‑point line for better readability. | Update existing workbooks to enforce a consistent divider appearance across all documents in a corporate template. | Iterate through multiple named line shapes in a worksheet to apply a standard dash style and weight as part of a branding guideline.
// AI Prompts: Generate C# code with Aspose.Cells that changes a shape's line color to red and sets a solid 1‑point line. | Show how to loop through all line shapes on a worksheet and assign a custom dash pattern and weight using Aspose.Cells. | Explain the steps to retrieve a shape by its name and modify its line properties (dash style, weight, color) in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

// The sample creates a workbook, inserts a line shape named "Divider", retrieves it by its identifier, changes the line's dash pattern to a dash style and sets the weight to 2 points, then writes the result to Output.xlsx using the Aspose.Cells C# API.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a line shape and give it the name "Divider"
            // Parameters: upper left row, upper left column, lower right row, lower right column, width, height
            Shape divider = sheet.Shapes.AddLine(1, 0, 5, 0, 200, 0);
            divider.Name = "Divider";

            // Retrieve the shape by its name
            Shape shape = sheet.Shapes["Divider"];

            // Change the line style to dashed
            shape.Line.DashStyle = MsoLineDashStyle.Dash;

            // Set the line thickness to 2 points
            shape.Line.Weight = 2.0;

            // Save the workbook
            workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
