// Title: Auto‑fit columns C‑F in Excel with Aspose.Cells for .NET (C#)
// Description: The example builds a workbook, populates cells C1‑F1 with sample text, and calls Worksheet.AutoFitColumns(2,5) to adjust the width of the selected columns before saving the file.
// Keywords: Aspose.Cells | AutoFitColumns | C# | .NET Excel automation | column width adjustment | specific column range | Worksheet.AutoFitColumns example | C to F columns
// Common Searches: Aspose.Cells auto fit specific columns C to F | Worksheet.AutoFitColumns usage with start and end indexes | C# code to auto size columns in Excel using Aspose.Cells | How to adjust column width by index in Aspose.Cells .NET | Excel column auto‑fit range Aspose.Cells
// Developer Intent: Resize columns C‑F so their contents fit without manual width settings.
// Use Cases: Generating financial reports where only columns C‑F need optimal width after data insertion. | Creating Excel templates that automatically size selected columns based on dynamic text lengths. | Building server‑side .NET services that export data and ensure proper column sizing for readability. | Designing dashboards where columns C‑F contain variable‑length strings and must be auto‑adjusted.
// AI Prompts: Provide a C# snippet that auto‑fits columns D‑H in an existing worksheet using Aspose.Cells. | Show how to auto‑fit columns by their letter names (e.g., "A" to "E") instead of numeric indexes with Aspose.Cells for .NET. | Explain the difference between Worksheet.AutoFitColumns() and Worksheet.AutoFitRows() in Aspose.Cells. | Create a reusable method that accepts start and end column letters and calls AutoFitColumns accordingly.

using System;
using Aspose.Cells;

namespace AutoFitColumnsExample
{
    // The example builds a workbook, populates cells C1‑F1 with sample text, and calls Worksheet.AutoFitColumns(2,5) to adjust the width of the selected columns before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data in columns C to F (indices 2 to 5)
            worksheet.Cells["C1"].PutValue("Short");
            worksheet.Cells["D1"].PutValue("Medium length text");
            worksheet.Cells["E1"].PutValue("This is a considerably longer piece of text that should cause the column to expand");
            worksheet.Cells["F1"].PutValue("Another long text example for column F");

            // Auto‑fit columns C (index 2) through F (index 5)
            worksheet.AutoFitColumns(2, 5);

            // Save the workbook
            workbook.Save("AutoFitColumns_C_to_F.xlsx");
        }
    }
}
