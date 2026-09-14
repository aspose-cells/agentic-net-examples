// Title: How to format a worksheet cell as currency with two decimal places using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that sets a cell's style to the built‑in currency format (Number = 7) and saves the workbook with Aspose.Cells. | Show an example of applying a currency style with two decimal places to a range of cells in an Excel file using Aspose.Cells for .NET. | Generate a snippet that retrieves a cell's current style, changes its number format to display currency, and writes the result to a .xlsx file with Aspose.Cells.
// Common Searches: Aspose.Cells C# set cell to display as currency with two decimal places | How to format Excel values as USD currency using Aspose.Cells API | C# example for applying a currency number format to an entire column in a workbook | Programmatically change numeric cell style to show currency in Aspose.Cells | Saving an Excel file with currency‑formatted cells using Aspose.Cells for .NET
// Tags: Aspose.Cells built‑in number format usage | apply currency cell style .NET | precision two decimals numeric formatting Aspose | adjust cell style numeric format C# | export workbook with currency formatting Aspose

using System;
using Aspose.Cells;

// // Demonstrates using Aspose.Cells for .NET to apply the built‑in currency number format (Number = 7) with two decimal places to a cell and save the workbook.
class CurrencyFormatDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Place a numeric value in cell A1
        Cell cell = sheet.Cells["A1"];
        cell.PutValue(1234.567);

        // Retrieve the current style of the cell
        Style style = cell.GetStyle();

        // Apply built‑in currency format with two decimal places (Number = 7)
        style.Number = 7;

        // Set the modified style back to the cell
        cell.SetStyle(style);

        // Save the workbook to a file
        workbook.Save("CurrencyFormatDemo.xlsx");
    }
}
