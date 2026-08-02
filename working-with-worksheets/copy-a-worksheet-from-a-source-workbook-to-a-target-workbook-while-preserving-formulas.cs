// Title: Copy Worksheet Between Workbooks with Formulas Preserved – Aspose.Cells for .NET (C#)
// Description: Demonstrates how to load a source workbook, select a worksheet, and copy it into a new workbook using Aspose.Cells' Worksheet.Copy method. The operation retains cell values, formatting, and all formulas, then saves the result as a separate file.
// Keywords: Aspose.Cells copy worksheet C# | preserve formulas Aspose.Cells | copy sheet to another workbook .NET | Worksheet.Copy example | duplicate Excel sheet programmatically
// Common Searches: Aspose.Cells copy worksheet keep formulas | C# copy Excel sheet to new workbook | Worksheet.Copy preserving formulas | How to duplicate a sheet with Aspose.Cells | Copy sheet from source.xlsx to output.xlsx C#
// Developer Intent: Programmatically duplicate a worksheet from one workbook to another while keeping all formulas and formatting intact.
// Use Cases: Generate client‑specific reports by copying a template sheet that contains calculation logic. | Create a master workbook and distribute individual sheets to separate projects without breaking formulas. | Automate the assembly of a final workbook by merging multiple pre‑formatted worksheets from different sources.
// AI Prompts: Write C# code using Aspose.Cells to copy several worksheets from a source workbook to a destination workbook, ensuring formulas remain functional. | Explain how to test that formulas were preserved after using Worksheet.Copy in Aspose.Cells. | Provide a step‑by‑step guide for copying a worksheet while converting the file format (e.g., XLSX to CSV) with Aspose.Cells.

using System;
using Aspose.Cells;

namespace WorksheetCopyExample
{
    // Demonstrates how to load a source workbook, select a worksheet, and copy it into a new workbook using Aspose.Cells' Worksheet.Copy method. The operation retains cell values, formatting, and all formulas, then saves the result as a separate file.
    class Program
    {
        static void Main()
        {
            // Load the source workbook (replace with your actual file path)
            Workbook sourceWorkbook = new Workbook("source.xlsx");

            // Get the worksheet you want to copy (e.g., the first worksheet)
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Create a new (empty) destination workbook
            Workbook destWorkbook = new Workbook();

            // Get the first worksheet of the destination workbook where the copy will be placed
            Worksheet destSheet = destWorkbook.Worksheets[0];

            // Copy the source worksheet into the destination worksheet.
            // This method copies contents, formats, and formulas, preserving the original formulas.
            destSheet.Copy(sourceSheet);

            // Save the destination workbook (replace with your desired output path)
            destWorkbook.Save("output.xlsx");
        }
    }
}
