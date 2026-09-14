// Title: Load an Excel workbook from a FileStream, change cell values, recalculate all formulas, and save as a new XLSX with Aspose.Cells for .NET
// AI Prompts: Read an existing XLSX file via FileStream, set cell A1 to 10, assign a formula to B1 that multiplies A1 by 2, invoke Workbook.CalculateFormula to update all dependent cells, and write the workbook to output.xlsx using Aspose.Cells in C#. | Using Aspose.Cells for .NET, open a workbook from a stream, modify cell contents and formulas, trigger full workbook recalculation, then save the updated file in XLSX format.
// Common Searches: Aspose.Cells C# load workbook from FileStream and recalculate formulas | How to update cell values and formulas then save workbook with Aspose.Cells | Programmatically trigger full formula calculation after editing cells in Aspose.Cells | Saving modified Excel file to a new location after calling Workbook.CalculateFormula in .NET
// Tags: load workbook from filestream Aspose.Cells | set cell value and formula C# | recalculate all formulas Workbook.CalculateFormula | save workbook as xlsx Aspose.Cells | modify worksheet cells programmatically

using System;
using System.IO;
using Aspose.Cells;

// The example opens input.xlsx via FileStream, loads it into a Workbook, sets A1 to 10, assigns B1 a formula '=A1*2', recalculates all formulas with Workbook.CalculateFormula, and saves the result as output.xlsx.
class Program
{
    static void Main()
    {
        // Open an existing Excel file as a stream
        using (FileStream fileStream = new FileStream("input.xlsx", FileMode.Open, FileAccess.Read))
        {
            // Load the workbook from the stream (Workbook(Stream) constructor)
            Workbook workbook = new Workbook(fileStream);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Modify some cells
            sheet.Cells["A1"].PutValue(10);          // Set a numeric value
            sheet.Cells["B1"].Formula = "=A1*2";     // Set a formula that depends on A1

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the updated workbook to a new file
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}
