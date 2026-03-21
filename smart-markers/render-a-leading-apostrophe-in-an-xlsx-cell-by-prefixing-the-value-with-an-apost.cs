using System;
using Aspose.Cells;

namespace AsposeCellsQuotePrefixDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the target cell (e.g., B10)
            Cell cell = worksheet.Cells["B10"];

            // Prefix the value with a single quote to indicate text
            cell.PutValue("'12345");

            // Retrieve the cell's style and enable QuotePrefix
            Style style = cell.GetStyle();
            style.QuotePrefix = true; // Marks the leading apostrophe as formatting, not part of the value

            // Apply the modified style back to the cell
            cell.SetStyle(style);

            // Save the workbook (lifecycle save)
            workbook.Save("QuotePrefixOutput.xlsx");
        }
    }
}