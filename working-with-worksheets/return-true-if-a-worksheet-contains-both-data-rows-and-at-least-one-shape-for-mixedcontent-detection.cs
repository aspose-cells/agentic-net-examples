using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

public class MixedContentDetector
{
    // Returns true if the worksheet has at least one non‑blank row and at least one shape.
    public static bool HasDataRowsAndShape(Worksheet worksheet)
    {
        // Determine if any row contains data.
        bool hasDataRow = false;
        int maxRow = worksheet.Cells.MaxDataRow; // last row with data
        for (int i = 0; i <= maxRow; i++)
        {
            Row row = worksheet.Cells.Rows[i];
            if (!row.IsBlank)
            {
                hasDataRow = true;
                break;
            }
        }

        // Determine if the worksheet contains any shapes.
        bool hasShape = worksheet.Shapes.Count > 0;

        return hasDataRow && hasShape;
    }

    // Example usage.
    public static void Main()
    {
        // Create a new workbook and get the first worksheet.
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data to create a non‑blank row.
        sheet.Cells["A1"].PutValue("Header");
        sheet.Cells["A2"].PutValue("Value");

        // Add a shape to the worksheet.
        sheet.Shapes.AddRectangle(1, 1, 0, 0, 100, 50);

        // Check for mixed content.
        bool result = HasDataRowsAndShape(sheet);
        Console.WriteLine("Worksheet has data rows and at least one shape: " + result);

        // Save the workbook (optional).
        workbook.Save("MixedContentDemo.xlsx");
    }
}