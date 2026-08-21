// Title: CopyRows to Duplicate Multiple Rows and Preserve Shared Formulas in Aspose.Cells for .NET
// Description: Demonstrates how to copy a block of rows (A1:B5) using Cells.CopyRows, automatically adjust shared formulas (e.g., =A1*2) to the new rows (A6:B10), verify the updated formulas, and save the workbook as an XLSX file.
// Keywords: Aspose.Cells CopyRows | duplicate rows C# | shared formula update | copy rows preserve formulas | Aspose.Cells .NET example | Excel automation C#
// Common Searches: Aspose.Cells copy rows with formulas | C# copy multiple rows in Excel workbook | preserve shared formulas when duplicating rows | CopyRows method example Aspose.Cells | verify formula references after row copy
// Developer Intent: Programmatically duplicate a range of rows and have all formulas automatically reference the new row positions.
// Use Cases: Clone a data block that contains calculations to create a new section with correct references. | Copy template rows (headers, formulas) for batch data entry without manual formula adjustments. | Automated testing to ensure formula integrity after row duplication in generated reports.
// AI Prompts: Show how to use Aspose.Cells CopyRows in C# to duplicate rows while keeping shared formulas correct. | Provide code that copies rows 1‑5 to rows 6‑10 and prints the updated formulas for verification. | Explain the behavior of shared formulas when rows are copied with Cells.CopyRows and how to validate them.

using System;
using Aspose.Cells;

// Demonstrates how to copy a block of rows (A1:B5) using Cells.CopyRows, automatically adjust shared formulas (e.g., =A1*2) to the new rows (A6:B10), verify the updated formulas, and save the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate column A with values 1 to 5
        for (int i = 0; i < 5; i++)
        {
            cells[i, 0].PutValue(i + 1); // A1..A5
        }

        // Set a shared formula in column B that multiplies column A by 2
        // This will automatically fill the formula down for the next 4 rows (total 5 rows)
        cells[0, 1].SetSharedFormula("=A1*2", 5, 1);

        // Duplicate the first 5 rows (0‑based index) and insert them starting at row index 5
        // After this operation rows 6‑10 will be copies of rows 1‑5
        cells.CopyRows(cells, 0, 5, 5);

        // Verify that the formulas in the duplicated rows have been updated correctly
        // Expected: B6 =A6*2, B7 =A7*2, ..., B10 =A10*2
        for (int row = 5; row < 10; row++)
        {
            Console.WriteLine($"Cell B{row + 1} formula: {cells[row, 1].Formula}");
        }

        // Save the workbook to a file
        workbook.Save("CopyRowsDuplicate.xlsx");
    }
}
