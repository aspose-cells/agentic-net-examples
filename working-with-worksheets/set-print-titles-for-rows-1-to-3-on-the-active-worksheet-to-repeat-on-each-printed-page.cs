// Title: Set rows 1‑3 as repeating print titles in an Aspose.Cells .NET worksheet (C#)
// Description: Creates a workbook, adds sample data, and assigns "$1:$3" to Worksheet.PageSetup.PrintTitleRows so rows 1‑3 repeat on every printed page, then saves the file as PrintTitleRowsDemo.xlsx.
// Keywords: Aspose.Cells | C# | .NET | PrintTitleRows | repeat header rows | set print titles | worksheet page setup | Excel printing | Aspose.Cells API example
// Common Searches: Aspose.Cells set PrintTitleRows C# | repeat first three rows on each printed page Aspose.Cells | how to define print title rows in .NET Excel library | Aspose.Cells worksheet page setup repeat rows | C# code to set print titles in Excel file
// Developer Intent: Configure rows 1‑3 to act as print titles so they appear on every printed page of the active worksheet.
// Use Cases: Printing a multi‑page report where the top three rows contain column headings that must be visible on each page. | Generating invoices where the first three rows hold company logo and address information that should repeat on every printed sheet. | Preparing a large data table for hard‑copy distribution, ensuring header rows stay at the top of each printed page.
// AI Prompts: Show how to set non‑contiguous rows as print titles with Aspose.Cells for .NET. | Provide C# code that sets PrintTitleRows, defines a print area, and changes page orientation using Aspose.Cells. | Explain how to programmatically verify that PrintTitleRows is applied when the workbook is opened in Excel.

using System;
using Aspose.Cells;

namespace AsposeCellsPrintTitleRowsDemo
{
    // Creates a workbook, adds sample data, and assigns "$1:$3" to Worksheet.PageSetup.PrintTitleRows so rows 1‑3 repeat on every printed page, then saves the file as PrintTitleRowsDemo.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();

            // Access the first (active) worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Optionally add some data to demonstrate the effect
            worksheet.Cells["A1"].PutValue("Header Row 1");
            worksheet.Cells["A2"].PutValue("Header Row 2");
            worksheet.Cells["A3"].PutValue("Header Row 3");
            for (int i = 4; i <= 20; i++)
            {
                worksheet.Cells[$"A{i}"].PutValue($"Data Row {i}");
            }

            // Set rows 1 to 3 as print titles (repeat on each printed page)
            // Using the PageSetup.PrintTitleRows property as defined in the API
            worksheet.PageSetup.PrintTitleRows = "$1:$3";

            // Save the workbook (lifecycle save rule)
            workbook.Save("PrintTitleRowsDemo.xlsx");

            Console.WriteLine("Workbook saved with print title rows set to $1:$3.");
        }
    }
}
