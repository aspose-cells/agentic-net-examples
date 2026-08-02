// Title: Detect data rows and shapes in an Aspose.Cells worksheet (C#)
// Description: Demonstrates using Aspose.Cells for .NET to verify that a worksheet contains at least one non‑blank row and one drawing shape, returning true for mixed content.
// Keywords: Aspose.Cells | C# | .NET | worksheet data detection | shape detection | mixed content | Excel rows | drawing objects | worksheet.Shapes | Excel automation
// Common Searches: Aspose.Cells check if worksheet has data and shapes | C# detect non‑blank rows in Aspose.Cells | How to find shapes in Aspose.Cells worksheet | Mixed content detection Aspose.Cells .NET | Determine if Excel sheet contains graphics and data using Aspose
// Developer Intent: Identify worksheets that contain at least one data row and at least one drawing shape.
// Use Cases: Skip empty or graphics‑only sheets during bulk report generation. | Validate worksheets before PDF export to ensure both tables and graphics render correctly. | Apply watermarks or custom logic only on sheets that have both data and shapes. | Optimize processing pipelines by handling mixed‑content sheets separately.
// AI Prompts: Write NUnit unit tests for HasDataAndShape covering sheets with only data, only shapes, both, and neither. | Refactor the method to use worksheet.Cells.MaxDataRow and early exit for better performance. | Generate C# code that returns false when a worksheet contains shapes but no data rows, using Aspose.Cells API. | Create a PowerShell script that invokes HasDataAndShape via .NET Core for batch workbook analysis.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates using Aspose.Cells for .NET to verify that a worksheet contains at least one non‑blank row and one drawing shape, returning true for mixed content.
public class MixedContentDetector
{
    // Returns true if the worksheet has at least one non‑blank row and at least one shape.
    public static bool HasDataAndShape(Worksheet worksheet)
    {
        // Determine whether any row contains data.
        bool hasData = false;
        foreach (Row row in worksheet.Cells.Rows)
        {
            if (!row.IsBlank)
            {
                hasData = true;
                break;
            }
        }

        // Determine whether the worksheet contains any drawing shapes.
        bool hasShape = worksheet.Shapes.Count > 0;

        return hasData && hasShape;
    }

    public static void Main()
    {
        // Create a new workbook.
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data.
        sheet.Cells["A1"].PutValue("Header");
        sheet.Cells["A2"].PutValue(123);

        // Add a shape (rectangle) to the worksheet.
        sheet.Shapes.AddRectangle(1, 1, 0, 0, 100, 50);

        // Check for mixed content.
        bool mixedContent = HasDataAndShape(sheet);
        Console.WriteLine("Worksheet contains data rows and at least one shape: " + mixedContent);
    }
}
