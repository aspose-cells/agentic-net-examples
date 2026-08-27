// Title: How to set a fixed column width and auto‑fit another column in an Aspose.Cells .NET workbook (C#)
// AI Prompts: Write C# code that creates a new workbook, adds short text to column A and long text to column B, sets column A to 15 characters with the SetColumnWidth method, and then applies AutoFitColumn to column B before saving as an .xlsx file. | Show a side‑by‑side comparison of a manually defined column width versus an automatically adjusted column by using SetColumnWidth for one column and AutoFitColumn for the adjacent column in Aspose.Cells.
// Common Searches: Aspose.Cells C# specify column width in characters | C# adjust column width automatically from cell values using Aspose.Cells | example of manual column width versus auto‑adjusted column in .NET Excel library | how to compare SetColumnWidth and AutoFitColumn results in an Aspose.Cells worksheet
// Tags: SetColumnWidth method Aspose.Cells C# | AutoFitColumn usage Aspose.Cells .xlsx | manual vs auto column sizing Aspose.Cells | column width comparison workbook Aspose.Cells | programmatic column sizing .NET Excel

using System;
using Aspose.Cells;

namespace ColumnWidthDemo
{
    // The program creates a workbook, writes short strings to column A and long strings to column B, sets column A width to 15 characters with SetColumnWidth, auto‑fits column B with AutoFitColumn, and saves the file as ColumnWidthComparison.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate column A with short text
            cells["A1"].PutValue("Short");
            cells["A2"].PutValue("Data");
            cells["A3"].PutValue("Info");

            // Populate column B with longer text that will require auto‑fit
            cells["B1"].PutValue("This is a much longer piece of text that should cause the column to expand when auto‑fit is applied.");
            cells["B2"].PutValue("Another lengthy entry to demonstrate the effect of AutoFitColumn.");
            cells["B3"].PutValue("Yet another long string for testing purposes.");

            // Set a fixed width for column A (index 0) – 15 characters
            cells.SetColumnWidth(0, 15.0);

            // Auto‑fit column B (index 1) based on its content
            sheet.AutoFitColumn(1);

            // Save the workbook
            workbook.Save("ColumnWidthComparison.xlsx");
        }
    }
}
