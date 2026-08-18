// Title: Aspose.Cells for .NET: Set Row Height and Auto‑Fit the Next Row with Wrapped Text (C#)
// Description: This C# example creates a workbook, writes short text to A1 and a long wrapped string to A2, enables text wrapping, sets the first row height to 20 points with SetRowHeight, auto‑fits the second row using AutoFitRow, and saves the file as RowHeightAutoFitDemo.xlsx.
// Keywords: Aspose.Cells | C# | .NET | set row height | AutoFitRow | wrap text | Excel row height programmatically | Aspose.Cells tutorial | Excel automation | GitHub example
// Common Searches: How to set a specific row height in Aspose.Cells C# | Aspose.Cells AutoFitRow with wrapped text example | C# code to programmatically adjust Excel row height using Aspose.Cells | Set custom row height then auto‑fit another row Aspose.Cells
// Developer Intent: Define a fixed height for one row and let the following row automatically resize to fit its wrapped content.
// Use Cases: Design a report where the header row keeps a constant height while data rows expand for multi‑line descriptions. | Generate invoices where the title row stays uniform but each item description row adapts to varying text lengths.
// AI Prompts: Generate C# code with Aspose.Cells that sets row 0 height to 25 points, enables text wrapping for a cell in row 1, and then auto‑fits row 1. | Explain the interaction between AutoFitRow and text wrapping in Aspose.Cells, including required style settings. | Provide a snippet that assigns custom heights to multiple rows and auto‑fits a selected row based on its content.

using System;
using Aspose.Cells;

namespace AsposeCellsRowHeightDemo
{
    // This C# example creates a workbook, writes short text to A1 and a long wrapped string to A2, enables text wrapping, sets the first row height to 20 points with SetRowHeight, auto‑fits the second row using AutoFitRow, and saves the file as RowHeightAutoFitDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate first row with short text
            worksheet.Cells["A1"].PutValue("Short text");
            // Populate second row with long wrapped text to demonstrate autofit
            worksheet.Cells["A2"].PutValue("This is a long piece of text that will require the row height to increase when auto‑fitted.");
            // Enable text wrapping for the long text cell
            Style wrapStyle = worksheet.Cells["A2"].GetStyle();
            wrapStyle.IsTextWrapped = true;
            worksheet.Cells["A2"].SetStyle(wrapStyle);

            // Set a custom height for the first row (index 0)
            worksheet.Cells.SetRowHeight(0, 20); // height in points

            // Auto‑fit the second row (index 1) based on its content
            worksheet.AutoFitRow(1); // uses the entire row range

            // Save the workbook to a file
            workbook.Save("RowHeightAutoFitDemo.xlsx");
        }
    }
}
