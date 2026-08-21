// Title: C# Example: Audit Shape Adjustment Guide Changes to a Text File with Aspose.Cells
// Description: Shows how to create a workbook, add a RightArrowCallout auto shape, modify its ShapeGuideCollection, and write each adjustment change—including added guides—to a timestamped text log before saving the file. Perfect for compliance‑oriented auditing of shape geometry in .NET.
// Keywords: Aspose.Cells | C# shape adjustment | ShapeGuideCollection | audit log | record shape changes | Excel shape guide modification | auto shape adjustment | timestamped log | compliance reporting | Aspose.Cells .NET example
// Common Searches: Aspose.Cells log shape guide changes | C# audit shape adjustments in Excel | write shape adjustment values to file Aspose | track auto shape geometry modifications .NET | shape adjustment guide audit example
// Developer Intent: Create a text‑based audit trail of every shape adjustment guide modification (including additions) while working with Aspose.Cells in C#.
// Use Cases: Generate compliance reports that capture every change to shape geometry before workbook distribution. | Debug and verify auto‑shape adjustments during automated Excel report generation. | Maintain a version‑controlled log of shape guide values for collaborative design workflows. | Provide a QA audit log for regulatory or internal review of workbook design changes.
// AI Prompts: Provide C# code that logs shape guide changes to a CSV file using Aspose.Cells. | Show how to encapsulate the shape adjustment logging into a reusable class with a configurable log file path. | Explain how to capture original and new guide values and output them as a JSON audit file. | Create a PowerShell script that runs the compiled example and archives the generated log file.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to create a workbook, add a RightArrowCallout auto shape, modify its ShapeGuideCollection, and write each adjustment change—including added guides—to a timestamped text log before saving the file. Perfect for compliance‑oriented auditing of shape geometry in .NET.
class ShapeAdjustmentAudit
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add an auto shape that supports adjustment guides
        Shape shape = worksheet.Shapes.AddAutoShape(
            AutoShapeType.RightArrowCallout, // shape type
            2, 0, 2, 0,                     // upper-left row, column, offsetX, offsetY
            200, 150);                      // width, height

        // Access the collection of shape adjustment guides
        ShapeGuideCollection guides = shape.Geometry.ShapeAdjustValues;

        // Path to the audit log file
        string auditFilePath = "ShapeAdjustmentAudit.txt";

        // Open the audit file for writing
        using (StreamWriter writer = new StreamWriter(auditFilePath, false))
        {
            writer.WriteLine($"Audit started: {DateTime.Now}");

            // Iterate through existing guides, modify them, and log each change
            for (int i = 0; i < guides.Count; i++)
            {
                double oldValue = guides[i].Value;          // capture old value
                double newValue = oldValue + 10;            // example modification
                guides[i].Value = newValue;                 // apply new value

                // Log the change
                writer.WriteLine($"Guide {i}: changed from {oldValue} to {newValue}");
            }

            // Add a new adjustment guide and log the addition
            int addedIndex = guides.Add("adjNew", 30.0);
            writer.WriteLine($"Added new guide at index {addedIndex} with value 30.0");
        }

        // Save the workbook with the modified shape
        workbook.Save("ShapeAdjustmentAuditDemo.xlsx");
    }
}
