// Title: Apply Accounting Number Format to a Cell with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, inserts a numeric total‑sales value into B2, retrieves the cell style, sets the built‑in accounting format (Number = 37, pattern "#,##0_);(#,##0)"), reapplies the style, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells | C# accounting format | Excel accounting number format | style.Number 37 | format total sales cell | Aspose.Cells .NET | financial spreadsheet formatting
// Common Searches: Aspose.Cells apply accounting format C# | style.Number 37 accounting pattern Aspose | format cell as accounting in .NET Excel library | set accounting number format for a single cell using Aspose.Cells | C# code to display negative numbers in parentheses with Aspose
// Developer Intent: Use Aspose.Cells for .NET to apply the built‑in accounting number format to a specific Excel cell.
// Use Cases: Present total‑sales figures in financial reports with standard accounting styling. | Generate invoices where negative amounts are shown in parentheses. | Maintain consistent currency formatting across auto‑generated spreadsheets.
// AI Prompts: Provide C# code to apply the accounting format to an entire column with Aspose.Cells. | Show how to customize the accounting format pattern beyond the built‑in style 37. | Explain how to combine accounting number format with conditional formatting in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsAccountingFormatDemo
{
    // Creates a workbook, inserts a numeric total‑sales value into B2, retrieves the cell style, sets the built‑in accounting format (Number = 37, pattern "#,##0_);(#,##0)"), reapplies the style, and saves the file as an Excel workbook.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a numeric total sales value into cell B2
            Cell totalSalesCell = sheet.Cells["B2"];
            totalSalesCell.PutValue(123456.78);

            // Retrieve the cell's current style
            Style style = totalSalesCell.GetStyle();

            // Apply the Accounting number format (value 37 corresponds to "#,##0_);(#,##0)")
            style.Number = 37;

            // Assign the modified style back to the cell
            totalSalesCell.SetStyle(style);

            // Save the workbook to a file
            workbook.Save("TotalSales_AccountingFormat.xlsx");

            Console.WriteLine("Accounting format applied to total sales cell and workbook saved.");
        }
    }
}
