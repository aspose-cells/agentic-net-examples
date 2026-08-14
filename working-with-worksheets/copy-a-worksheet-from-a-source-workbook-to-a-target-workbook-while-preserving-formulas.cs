// Title: Copy a Worksheet to a New Workbook with Formulas Preserved – Aspose.Cells C# Example
// Description: Demonstrates loading a source workbook, copying its first worksheet into a new workbook using Aspose.Cells' Worksheet.Copy method, retaining all cell values, formats, and formulas, and saving the result as an XLSX file.
// Keywords: Aspose.Cells | C# worksheet copy | preserve formulas | copy worksheet to new workbook | Worksheet.Copy | Excel automation .NET | duplicate sheet with formulas | Aspose.Cells example
// Common Searches: Aspose.Cells copy worksheet preserve formulas | C# copy Excel sheet to new file | Worksheet.Copy method example | How to duplicate a sheet with formulas in .NET | Copy sheet between workbooks Aspose.Cells
// Developer Intent: Copy a worksheet from one workbook to another while keeping all formulas intact.
// Use Cases: Create per‑client reports by cloning a template sheet into separate workbooks. | Migrate legacy spreadsheet data to a fresh file without breaking calculations. | Archive a calculation sheet by duplicating it into a standalone workbook. | Generate batch workbooks for data analysis while preserving formula logic.
// AI Prompts: Provide C# code using Aspose.Cells to copy a worksheet from a source workbook to a destination workbook, ensuring formulas are retained. | Show how to copy multiple worksheets with formulas and adjust external references in the target workbook. | Explain how to copy a sheet with formulas and then rename or update named ranges after the copy.

using System;
using Aspose.Cells;

namespace WorksheetCopyExample
{
    // Demonstrates loading a source workbook, copying its first worksheet into a new workbook using Aspose.Cells' Worksheet.Copy method, retaining all cell values, formats, and formulas, and saving the result as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Load the source workbook (replace with your actual file path)
            Workbook sourceWorkbook = new Workbook("source.xlsx");

            // Get the worksheet you want to copy (first worksheet in this example)
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Create a new (empty) destination workbook
            Workbook destWorkbook = new Workbook();

            // Get the first worksheet of the destination workbook where the copy will be placed
            Worksheet destSheet = destWorkbook.Worksheets[0];

            // Copy the source worksheet to the destination worksheet.
            // This method copies cells, formats, and formulas, preserving the original formulas.
            destSheet.Copy(sourceSheet);

            // Save the destination workbook (replace with your desired output path)
            destWorkbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}
