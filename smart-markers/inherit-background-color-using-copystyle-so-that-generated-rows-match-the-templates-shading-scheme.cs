// Title: Copy background shading from a template row to newly inserted rows with Aspose.Cells in C#
// AI Prompts: Insert a row and duplicate the source row’s background color using Row.CopySettings and ApplyStyle in Aspose.Cells C#. | Transfer a template row’s shading to generated rows by copying its style flags with Aspose.Cells. | Apply the CellShading style flag from a source row to a destination row after inserting rows in a workbook.
// Common Searches: Aspose.Cells how to copy row shading when inserting rows in C# | C# copy background color from one row to another using Aspose.Cells | Use CopySettings to preserve row formatting in Aspose.Cells workbook | Apply template row style to new rows Aspose.Cells C# example | Transfer cell shading between rows with Aspose.Cells API
// Tags: row.CopySettings with style flag cell shading | inherit row background color Aspose.Cells C# | apply template row style to inserted rows | transfer row formatting using Aspose.Cells | copy row shading Aspose.Cells workbook

using System;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsRowStyleCopyDemo
{
    // The code loads a template workbook, inserts a new row, copies the source row’s settings with Row.CopySettings (including style), then applies the source style using a StyleFlag that enables CellShading, and saves the workbook so the inserted row inherits the template’s background shading.
    class Program
    {
        static void Main()
        {
            // Load the template workbook that contains the desired row shading scheme
            Workbook workbook = new Workbook("Template.xlsx"); // load
            Worksheet sheet = workbook.Worksheets[0];

            // Define source row (the row whose background color we want to inherit)
            int sourceRowIndex = 0; // e.g., first row in the template
            Row sourceRow = sheet.Cells.Rows[sourceRowIndex];

            // Insert a new blank row at the desired position
            int destinationRowIndex = 5; // example target row index
            sheet.Cells.InsertRows(destinationRowIndex, 1);

            // Get the newly inserted row
            Row destinationRow = sheet.Cells.Rows[destinationRowIndex];

            // Copy all settings (including style) from the source row to the destination row
            // checkStyle = true ensures style is gathered correctly even if workbooks differ
            destinationRow.CopySettings(sourceRow, true);

            // Optionally, apply the style explicitly to ensure cell shading is transferred
            // This applies the style flags that affect cell shading
            Style sourceStyle = sourceRow.GetStyle();
            destinationRow.ApplyStyle(sourceStyle, new StyleFlag { CellShading = true });

            // Save the modified workbook
            workbook.Save("Result.xlsx"); // save
        }
    }
}
