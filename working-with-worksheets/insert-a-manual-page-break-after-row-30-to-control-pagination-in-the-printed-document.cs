// Title: Aspose.Cells for .NET – Insert a Manual Horizontal Page Break After Row 30
// Description: Creates a new workbook, fills column A with 60 rows, adds a manual horizontal page break at zero‑based index 30 (Excel row 31) to control printed pagination, and saves the file as ManualPageBreakAfterRow30.xlsx.
// Keywords: Aspose.Cells page break .NET | C# horizontal page break Excel | manual page break after row 30 | Excel pagination Aspose | worksheet.HorizontalPageBreaks.Add
// Common Searches: add manual page break Aspose.Cells C# | horizontal page break after specific row .NET | control Excel print pagination with Aspose | Aspose.Cells insert page break row 30 example
// Developer Intent: Add a manual horizontal page break after row 30 to manage printed page layout.
// Use Cases: Produce multi‑page reports where each page starts after a fixed number of rows. | Generate printable invoices that begin on a new sheet page at a defined row. | Export large data sets with consistent page breaks for accurate printing.
// AI Prompts: Show how to add horizontal page breaks at rows 20, 40, and 60 using Aspose.Cells for .NET. | Explain the steps to delete a specific horizontal page break from a worksheet. | Provide code that combines vertical page breaks with print options and horizontal breaks.

using System;
using Aspose.Cells;

namespace AsposeCellsPageBreakDemo
{
    // Creates a new workbook, fills column A with 60 rows, adds a manual horizontal page break at zero‑based index 30 (Excel row 31) to control printed pagination, and saves the file as ManualPageBreakAfterRow30.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data (optional, just to visualize the break)
            for (int i = 0; i < 60; i++)
            {
                worksheet.Cells[i, 0].PutValue($"Row {i + 1}");
            }

            // Insert a manual horizontal page break after row 30.
            // Row index is zero‑based, so row 30 corresponds to Excel row 31.
            worksheet.HorizontalPageBreaks.Add(30);

            // Save the workbook to an XLSX file
            workbook.Save("ManualPageBreakAfterRow30.xlsx");
        }
    }
}
