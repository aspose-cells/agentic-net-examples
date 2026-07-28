// Title: C# – Apply Accounting Number Format to a Cell with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, insert a total‑sales value, set the built‑in accounting format (Number ID 37, "#,##0_);(#,##0)") on cell B2, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells accounting format C# | Excel accounting number format .NET | Number format ID 37 Aspose | apply accounting style to cell | format monetary values Aspose.Cells
// Common Searches: Aspose.Cells set accounting number format C# | How to use Number = 37 for accounting style | Apply accounting format to Excel cell with Aspose | C# code example for accounting number format in Aspose.Cells | Save workbook after formatting cell as accounting
// Developer Intent: The developer needs to format a specific Excel cell with the accounting number style using Aspose.Cells for .NET.
// Use Cases: Prepare financial statements where totals are displayed in accounting format. | Standardize currency columns across multiple worksheets before exporting reports. | Create a helper method that applies accounting formatting to any numeric range based on user settings.
// AI Prompts: Generate C# code that applies the accounting number format (ID 37) to a range of cells with Aspose.Cells. | Write a reusable function to detect monetary cells and set the accounting style automatically. | Explain how to retrieve and customize the built‑in accounting format for different locales in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsAccountingFormatDemo
{
    // Demonstrates how to create a workbook, insert a total‑sales value, set the built‑in accounting format (Number ID 37, "#,##0_);(#,##0)") on cell B2, and save the file using Aspose.Cells for .NET.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a sample total sales value into cell B2
            Cell totalSalesCell = sheet.Cells["B2"];
            totalSalesCell.PutValue(123456.78);

            // Retrieve the current style of the cell
            Style style = totalSalesCell.GetStyle();

            // Apply the Accounting number format (value 37 corresponds to "#,##0_);(#,##0)")
            style.Number = 37;

            // Assign the modified style back to the cell
            totalSalesCell.SetStyle(style);

            // Save the workbook to a file
            workbook.Save("TotalSales_AccountingFormat.xlsx");

            Console.WriteLine("Accounting number format applied to cell B2 and workbook saved.");
        }
    }
}
