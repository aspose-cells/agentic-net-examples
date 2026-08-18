// Title: Set Header Row Height and Auto‑Fit Remaining Rows with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, write a header and data rows, apply a custom height to the first row, auto‑fit all subsequent rows, optionally auto‑fit columns, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# set row height | AutoFitRows Aspose.Cells | header row formatting .NET | adjust row height programmatically | AutoFitColumns Aspose.Cells | Excel row height C# | Aspose.Cells workbook example | row height and auto‑fit rows
// Common Searches: Aspose.Cells set header row height C# | auto fit rows after custom height Aspose.Cells | C# code to set row height and auto‑fit rows in Excel | how to use AutoFitRows range Aspose.Cells | Aspose.Cells example for row height and auto‑fit
// Developer Intent: Apply a fixed height to the header row while automatically resizing all other rows to fit their content.
// Use Cases: Designing reports where the title row needs a larger, consistent height and data rows adjust to varying text lengths. | Generating invoices with a bold header row and auto‑sized line‑item rows for readability. | Creating product catalogs that emphasize the header and automatically accommodate long descriptions in subsequent rows.
// AI Prompts: Show C# code that sets row 0 height to 30 points and auto‑fits rows 1 to the last data row using Aspose.Cells. | Explain how to combine SetRowHeight and AutoFitRows for a worksheet with a fixed header and dynamic data rows. | Provide an Aspose.Cells example that sets a custom header height, auto‑fits remaining rows, and optionally auto‑fits columns.

using System;
using Aspose.Cells;

namespace AsposeCellsHeaderRowHeightDemo
{
    // Demonstrates how to create a workbook, write a header and data rows, apply a custom height to the first row, auto‑fit all subsequent rows, optionally auto‑fit columns, and save the file using Aspose.Cells for .NET.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // -------------------------------------------------
            // 1. Populate some sample data (header + data rows)
            // -------------------------------------------------
            // Header row (row 0)
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Description");
            cells["C1"].PutValue("Price");

            // Data rows
            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue("Fresh red apple with a crisp texture");
            cells["C2"].PutValue(0.99);

            cells["A3"].PutValue("Banana");
            cells["B3"].PutValue("Ripe banana, perfect for smoothies and snacks");
            cells["C3"].PutValue(0.59);

            cells["A4"].PutValue("Cherry");
            cells["B4"].PutValue("Sweet cherries, great for desserts");
            cells["C4"].PutValue(2.49);

            // -------------------------------------------------
            // 2. Set a custom height for the header row (row 0)
            // -------------------------------------------------
            // Height is in points; 30 points gives a visually emphasized header
            sheet.Cells.SetRowHeight(0, 30);

            // -------------------------------------------------
            // 3. Auto‑fit the remaining rows (from row 1 to the last data row)
            // -------------------------------------------------
            int firstDataRow = 1;
            int lastDataRow = cells.MaxDataRow; // Gets the index of the last row that contains data
            if (lastDataRow >= firstDataRow)
            {
                // AutoFitRows(startRow, endRow) fits rows in the specified range
                sheet.AutoFitRows(firstDataRow, lastDataRow);
            }

            // -------------------------------------------------
            // 4. Optionally auto‑fit columns for better visibility
            // -------------------------------------------------
            sheet.AutoFitColumns();

            // -------------------------------------------------
            // 5. Save the workbook
            // -------------------------------------------------
            workbook.Save("HeaderRowHeightDemo.xlsx");
        }
    }
}
