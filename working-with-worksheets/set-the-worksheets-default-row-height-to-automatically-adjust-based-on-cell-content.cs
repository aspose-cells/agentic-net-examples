// Title: Auto‑fit worksheet row height to content with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, writes a long string to a cell, enables text wrapping, calls Worksheet.AutoFitRows() to let rows automatically resize to the wrapped content, and saves the file as XLSX.
// Keywords: Aspose.Cells AutoFitRows | C# auto row height | wrap text Aspose.Cells | default row height Excel .NET | adjust row height programmatically | Excel row height auto fit C#
// Common Searches: Aspose.Cells auto fit rows C# example | How to auto adjust row height in Excel using Aspose.Cells | Set default row height based on cell content Aspose | Enable text wrap and auto row height Aspose.Cells | C# code to auto‑fit rows Aspose.Cells
// Developer Intent: Programmatically make worksheet rows automatically resize to fit wrapped cell content.
// Use Cases: Generating reports where description fields vary in length and need automatic row expansion. | Creating invoices with notes that may span multiple lines without manual height tweaks. | Exporting data tables to Excel where cells contain multi‑line values and each row must adapt to its content.
// AI Prompts: Show C# code using Aspose.Cells to enable text wrapping for a cell and auto‑fit all rows so the row height matches the content. | Provide an Aspose.Cells example that sets the default row height automatically after inserting long text into several cells.

using System;
using System.IO;
using Aspose.Cells;

// Creates a workbook, writes a long string to a cell, enables text wrapping, calls Worksheet.AutoFitRows() to let rows automatically resize to the wrapped content, and saves the file as XLSX.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data that requires row height adjustment
        worksheet.Cells["A1"].PutValue("This is a long text that should cause the row height to increase automatically based on its content.");
        // Enable text wrapping so the content spans multiple lines
        Style style = worksheet.Cells["A1"].GetStyle();
        style.IsTextWrapped = true;
        worksheet.Cells["A1"].SetStyle(style);

        // Auto‑fit all rows in the worksheet; this adjusts the default row height
        // to match the content of the cells.
        worksheet.AutoFitRows();

        // Save the workbook to the desktop
        string outputPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "AutoFitRowsDefaultHeight.xlsx");
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}
