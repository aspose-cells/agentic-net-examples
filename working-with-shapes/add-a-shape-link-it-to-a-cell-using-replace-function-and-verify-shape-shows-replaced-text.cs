// Title: C# – Link a rectangle shape to a cell, replace placeholder text, and refresh the shape with Aspose.Cells for .NET
// Description: Demonstrates how to add a rectangle shape, link it to cell A1, use Workbook.Replace to swap a {{Name}} placeholder with actual text, and call UpdateSelectedValue so the shape shows the new value. The workbook is then saved as an XLSX file.
// Keywords: Aspose.Cells | C# | .NET | rectangle shape | SetLinkedCell | UpdateSelectedValue | Workbook.Replace | placeholder replacement | linked cell shape | dynamic shape text
// Common Searches: Aspose.Cells link shape to cell C# | Update shape text after Workbook.Replace | SetLinkedCell example Aspose.Cells | Refresh linked shape after placeholder substitution | C# Aspose.Cells replace placeholder in worksheet
// Developer Intent: Create a shape linked to a cell, replace a placeholder in that cell, and ensure the shape displays the updated content.
// Use Cases: Template generation where shape captions must reflect data inserted into linked cells. | Automated report creation that uses shapes as visual labels synchronized with cell values after bulk text replacement. | Dynamic dashboards where shape text updates automatically when underlying cell content changes.
// AI Prompts: Generate C# code using Aspose.Cells to add a rectangle shape, link it to A1, replace "{{Name}}" with a real name, and refresh the shape text. | Explain how SetLinkedCell and UpdateSelectedValue work together when cell content is modified via Workbook.Replace. | Suggest error‑handling patterns for linking shapes to cells and updating them after text substitution in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add a rectangle shape, link it to cell A1, use Workbook.Replace to swap a {{Name}} placeholder with actual text, and call UpdateSelectedValue so the shape shows the new value. The workbook is then saved as an XLSX file.
class ShapeReplaceDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Put a placeholder text into a cell that will be linked to the shape
        sheet.Cells["A1"].PutValue("{{Name}}");

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, upper left pixel X, upper left pixel Y, width, height
        Shape shape = sheet.Shapes.AddRectangle(2, 2, 100, 100, 0, 0);

        // Link the shape to the cell containing the placeholder
        shape.SetLinkedCell("$A$1", false, false);

        // Load the initial linked value into the shape
        shape.UpdateSelectedValue();

        // Replace the placeholder in the worksheet with the desired text
        workbook.Replace("{{Name}}", "John Doe");

        // Refresh the shape so it reflects the replaced value
        shape.UpdateSelectedValue();

        // Verify by printing the linked cell's current value (the shape displays this value)
        Console.WriteLine("Linked cell value after replace: " + sheet.Cells["A1"].StringValue);

        // Save the workbook (lifecycle rule)
        workbook.Save("ShapeReplaceDemo.xlsx");
    }
}
