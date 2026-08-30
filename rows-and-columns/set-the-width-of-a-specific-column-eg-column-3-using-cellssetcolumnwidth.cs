// Title: Set column C width to 25.5 characters with Aspose.Cells in C#
// AI Prompts: Write C# code that uses Aspose.Cells to set the width of column index 2 to 25.5 characters and save the workbook. | Demonstrate how to apply Cells.SetColumnWidth to adjust a specific column's width, insert a sample value, and export the file. | Create a minimal Aspose.Cells example that changes a column's width, puts text into C1, and writes the workbook to an .xlsx file.
// Common Searches: aspocells set column width c# example | how to change column width in Excel using Aspose.Cells .NET | C# Aspose.Cells SetColumnWidth column index 2 | adjust column width programmatically with Aspose.Cells library
// Tags: Aspose.Cells SetColumnWidth method | C# column width adjustment Excel | programmatic column sizing Aspose.Cells | Excel .xlsx column width C# | custom column width workbook Aspose

using System;
using Aspose.Cells;

namespace AsposeCellsColumnWidthExample
{
    // The program creates a new workbook, sets column C (zero‑based index 2) to a width of 25.5 characters using Cells.SetColumnWidth, writes a sample text into C1, and saves the file as ColumnWidthDemo.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (default contains one worksheet)
            Workbook workbook = new Workbook();

            // Access the first worksheet's cells collection
            Cells cells = workbook.Worksheets[0].Cells;

            // Set the width of column 3 (zero‑based index 2) to 25.5 characters
            cells.SetColumnWidth(2, 25.5);

            // Optional: put some data to visualize the column width
            cells["C1"].PutValue("Column 3 with custom width");

            // Save the workbook to a file
            workbook.Save("ColumnWidthDemo.xlsx");
        }
    }
}
