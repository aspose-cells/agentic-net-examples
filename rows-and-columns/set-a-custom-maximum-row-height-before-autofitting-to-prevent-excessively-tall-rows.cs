// Title: Cap Row Height with AutoFitterOptions (MaxRowHeight) in Aspose.Cells for .NET
// Description: Shows how to define a maximum row height before AutoFitRows by using AutoFitterOptions.MaxRowHeight and OnlyAuto, preventing wrapped text from producing overly tall rows while keeping manually set heights intact.
// Keywords: Aspose.Cells | .NET | AutoFitterOptions | MaxRowHeight | OnlyAuto | limit row height | auto fit rows | wrap text | Excel row height | prevent tall rows | C# example
// Common Searches: Aspose.Cells limit row height auto fit | AutoFitterOptions MaxRowHeight C# example | prevent rows from becoming too tall Aspose.Cells | OnlyAuto property usage Aspose.Cells | set maximum row height before AutoFitRows
// Developer Intent: The developer wants to restrict row height during auto‑fit to avoid excessively tall rows.
// Use Cases: Apply a ceiling to row height when auto‑fitting wrapped text in generated reports. | Auto‑fit only rows that still have the default height, leaving custom heights unchanged. | Create Excel files with long descriptions while maintaining a consistent layout.
// AI Prompts: How do I use AutoFitterOptions to set MaxRowHeight and OnlyAuto when auto‑fitting rows in Aspose.Cells for .NET? | Provide a C# snippet that caps row height at 40 points while auto‑fitting multiple rows. | Explain the effect of the OnlyAuto flag on rows with custom heights in Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to define a maximum row height before AutoFitRows by using AutoFitterOptions.MaxRowHeight and OnlyAuto, preventing wrapped text from producing overly tall rows while keeping manually set heights intact.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add long text and enable text wrapping to trigger row height increase
        worksheet.Cells["A1"].PutValue("This is a very long text that would normally cause the row to become excessively tall after auto‑fitting.");
        Style style = worksheet.Cells["A1"].GetStyle();
        style.IsTextWrapped = true;
        worksheet.Cells["A1"].SetStyle(style);

        // Optionally set an initial small row height
        worksheet.Cells.SetRowHeight(0, 10);

        // Configure AutoFitterOptions with a maximum row height
        AutoFitterOptions options = new AutoFitterOptions
        {
            MaxRowHeight = 50, // Limit maximum row height to 50 points
            OnlyAuto = true    // Apply auto‑fit only to rows without custom height
        };

        // Auto‑fit rows using the specified options
        worksheet.AutoFitRows(options);

        // Save the workbook
        workbook.Save("MaxRowHeightDemo.xlsx");
    }
}
