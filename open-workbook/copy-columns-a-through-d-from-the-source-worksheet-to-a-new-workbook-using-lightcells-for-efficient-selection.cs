// Title: C# – Copy Columns A‑D Between Workbooks Using Aspose.Cells LightCells
// Description: Loads a source workbook, creates a new workbook, and efficiently copies columns A through D (indices 0‑3) from the source worksheet to the destination worksheet with the LightCells‑enabled Cells.CopyColumns method, then saves the result.
// Keywords: Aspose.Cells | CopyColumns | LightCells | C# | .NET | Excel automation | copy columns A-D | workbook to workbook copy | efficient column copy | worksheet column transfer
// Common Searches: Aspose.Cells copy columns A to D C# | LightCells copy multiple columns .NET | How to transfer specific columns between Excel workbooks using Aspose.Cells | Efficient column copy with Cells.CopyColumns method | Copy column range from one worksheet to another Aspose.Cells
// Developer Intent: Copy the first four columns (A‑D) from a source worksheet into a new workbook with minimal overhead.
// Use Cases: Extract key columns from a large data sheet to create a lightweight summary workbook. | Generate a template that contains only required columns for downstream processing or reporting. | Share a subset of data with partners by copying selected columns into a separate file while preserving formatting.
// AI Prompts: Show a C# example that copies columns A‑D from one Excel file to another using Aspose.Cells LightCells and includes error handling. | Provide code that copies a column range and retains column widths and styles with Cells.CopyColumns in Aspose.Cells for .NET. | Explain how to use the Cells.CopyColumns method to move specific columns between workbooks and then save the destination file.

using System;
using Aspose.Cells;

namespace AsposeCellsCopyColumnsExample
{
    // Loads a source workbook, creates a new workbook, and efficiently copies columns A through D (indices 0‑3) from the source worksheet to the destination worksheet with the LightCells‑enabled Cells.CopyColumns method, then saves the result.
    class Program
    {
        static void Main()
        {
            // Load the source workbook (replace with your actual file path)
            Workbook sourceWorkbook = new Workbook("source.xlsx");
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Create a new workbook that will receive the copied columns
            Workbook destWorkbook = new Workbook();
            Worksheet destSheet = destWorkbook.Worksheets[0];

            // Copy columns A (index 0) through D (index 3) from the source sheet to the destination sheet
            // The fourth parameter specifies the number of columns to copy (4 columns: A, B, C, D)
            destSheet.Cells.CopyColumns(sourceSheet.Cells, 0, 0, 4);

            // Save the destination workbook
            destWorkbook.Save("dest.xlsx");
        }
    }
}
