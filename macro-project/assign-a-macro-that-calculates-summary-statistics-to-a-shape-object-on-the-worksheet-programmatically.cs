// Title: How to assign a VBA macro to a rectangle shape using Aspose.Cells for .NET (C#)
// Description: C# example that creates a workbook with Aspose.Cells, adds a rectangle shape to the first worksheet, sets its MacroName to "CalculateSummaryStats()", and saves the file as a macro‑enabled workbook, illustrating how to link a VBA macro to a shape programmatically.
// Keywords: Aspose.Cells C# macro shape | assign VBA macro to shape | Shape.MacroName Aspose.Cells | add rectangle shape .NET | link macro to Excel shape programmatically | CalculateSummaryStats macro | Excel automation Aspose.Cells
// Common Searches: Aspose.Cells assign macro to shape | C# set MacroName property on shape | link VBA macro to rectangle in Excel using .NET | programmatically attach macro to Excel shape | example of Shape.MacroName Aspose.Cells
// Developer Intent: Programmatically attach a VBA macro that calculates summary statistics to a shape in an Excel worksheet.
// Use Cases: Create a dashboard button that runs a summary‑statistics macro when clicked. | Enable one‑click report generation by linking a macro to a shape. | Add interactive chart annotations that trigger calculations via an attached macro.
// AI Prompts: Generate C# code with Aspose.Cells that adds a rectangle shape and assigns the macro "CalculateSummaryStats()" to it. | Show how to embed a VBA macro in a workbook and link it to a shape using Aspose.Cells for .NET. | Provide an example of retrieving and updating the MacroName of an existing shape in an Excel file with Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# example that creates a workbook with Aspose.Cells, adds a rectangle shape to the first worksheet, sets its MacroName to "CalculateSummaryStats()", and saves the file as a macro‑enabled workbook, illustrating how to link a VBA macro to a shape programmatically.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        Shape shape = worksheet.Shapes.AddRectangle(1, 1, 100, 100, 0, 0);

        // Assign a macro that calculates summary statistics to the shape
        shape.MacroName = "CalculateSummaryStats()";

        // Save the workbook
        workbook.Save("MacroShapeDemo.xlsx");
    }
}
