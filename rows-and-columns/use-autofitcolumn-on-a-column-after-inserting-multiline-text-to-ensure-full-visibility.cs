// Title: C# – AutoFit a column after inserting multiline wrapped text with Aspose.Cells
// Description: Creates a new workbook, writes multiline text to cell A1 using '\n', enables text wrapping, calls Worksheet.AutoFitColumn to size column A to the wrapped content, and saves the file as MultilineAutoFitColumn.xlsx.
// Keywords: Aspose.Cells AutoFitColumn C# | multiline text wrap Aspose.Cells | auto adjust column width .NET | Excel column autofit wrapped text | Aspose.Cells line break column width
// Common Searches: Aspose.Cells AutoFitColumn after line breaks | C# wrap text and autofit column in Excel | how to auto size column for wrapped text Aspose.Cells | auto fit column with multiline cell content .NET
// Developer Intent: Automatically resize a column so that multiline, wrapped text is fully visible in the generated Excel file.
// Use Cases: Reports with bullet‑point lists where each bullet is on a new line. | Invoices that contain multi‑line address fields requiring proper column width. | Product catalogs exporting paragraph‑style descriptions without truncation.
// AI Prompts: Provide C# code that inserts multiline text into a cell, enables wrapping, and uses AutoFitColumn to fit the column width with Aspose.Cells. | Explain how to auto‑fit multiple columns that contain wrapped text, including the row‑range parameters for Worksheet.AutoFitColumn. | Show a step‑by‑step example of adjusting column width after adding line breaks to a cell using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsAutoFitColumnExample
{
    // Creates a new workbook, writes multiline text to cell A1 using '\n', enables text wrapping, calls Worksheet.AutoFitColumn to size column A to the wrapped content, and saves the file as MultilineAutoFitColumn.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Insert multiline text into cell A1
            // Use '\n' to create line breaks
            worksheet.Cells["A1"].PutValue("First line\nSecond line\nThird line");

            // Enable text wrapping for the cell so that the multiline text is displayed correctly
            Style style = worksheet.Cells["A1"].GetStyle();
            style.IsTextWrapped = true;
            worksheet.Cells["A1"].SetStyle(style);

            // AutoFit the column that contains the multiline text (column A, index 0)
            // Specify the row range that contains data (row 0 to 0 in this case)
            worksheet.AutoFitColumn(0, 0, 0);

            // Save the workbook (lifecycle rule: save)
            workbook.Save("MultilineAutoFitColumn.xlsx");

            Console.WriteLine("Workbook saved with AutoFitColumn applied.");
        }
    }
}
