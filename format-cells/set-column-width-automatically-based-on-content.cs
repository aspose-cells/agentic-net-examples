// Title: Auto‑Fit Excel Column Widths with Aspose.Cells for .NET (C#)
// Description: Demonstrates creating a workbook, inserting short, medium, and long text plus a numeric value, invoking Worksheet.AutoFitColumns() to size each column to its longest cell, and saving the file as AutoFitColumnsResult.xlsx.
// Keywords: Aspose.Cells AutoFitColumns C# | auto fit column width Aspose.Cells | C# set Excel column width automatically | Worksheet.AutoFitColumns example | auto‑size Excel columns .NET | Aspose.Cells column width based on content | fit columns to text Aspose | auto adjust Excel column width C# | Aspose.Cells column autosizing | Excel column auto‑fit using Aspose
// Common Searches: How to auto‑fit columns in an Aspose.Cells workbook using C# | C# code to automatically adjust Excel column width with Aspose.Cells | Aspose.Cells AutoFitColumns method usage | Set column width dynamically based on cell values in Aspose.Cells for .NET | Auto‑size specific column range Aspose.Cells C# | Maximum column width with AutoFitColumns Aspose.Cells | Auto‑fit columns after merging cells Aspose.Cells C#
// Developer Intent: Automatically resize worksheet columns so each fits the longest content in its column.
// Use Cases: Create printable Excel reports where column widths adapt to content | Export DataGrid or DataTable to Excel with columns auto‑sized for readability | Generate Excel files for end‑users without manual formatting | Prepare spreadsheets for API consumers ensuring no truncated data
// AI Prompts: Show C# code to auto‑fit columns for a selected range (e.g., A1:C10) using Aspose.Cells. | Provide an example that auto‑fits columns after applying cell styles and merged cells in C#. | Explain how to limit the maximum column width while still using AutoFitColumns in Aspose.Cells. | Give a snippet to auto‑fit columns based on content after inserting images in a worksheet.

using System;
using Aspose.Cells;

namespace AsposeCellsAutoFitColumnDemo
{
    // Demonstrates creating a workbook, inserting short, medium, and long text plus a numeric value, invoking Worksheet.AutoFitColumns() to size each column to its longest cell, and saving the file as AutoFitColumnsResult.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate cells with varying length text to demonstrate auto‑fit
            worksheet.Cells["A1"].PutValue("Short");
            worksheet.Cells["B1"].PutValue("Medium length text");
            worksheet.Cells["C1"].PutValue("This is a much longer piece of text that should cause the column to expand automatically");
            worksheet.Cells["D1"].PutValue(12345.6789);
            worksheet.Cells["E1"].PutValue("Another long text entry to test the auto‑fit functionality");

            // Auto‑fit all columns based on the content in the worksheet
            worksheet.AutoFitColumns();

            // Save the workbook to a file
            workbook.Save("AutoFitColumnsResult.xlsx");
        }
    }
}
