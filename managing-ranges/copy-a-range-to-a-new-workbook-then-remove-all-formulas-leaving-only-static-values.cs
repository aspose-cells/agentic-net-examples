// Title: Copy a Range to a New Workbook and Convert Formulas to Static Values with Aspose.Cells for .NET
// Description: Creates a source workbook, adds numeric data and a SUM formula, defines the source range A1:A3, creates a destination workbook, copies the range with CopyData, removes all formulas using Worksheet.Cells.RemoveFormulas, and saves the result as a static‑value workbook (CopiedValues.xlsx).
// Keywords: Aspose.Cells | CopyData | RemoveFormulas | C# copy range | static values | export workbook | formula to value | range copy .NET | Excel snapshot | no formulas
// Common Searches: Aspose.Cells copy range to another workbook | remove formulas after copying range Aspose | convert Excel formulas to values C# | how to export static values with Aspose.Cells | CopyData vs PasteValues Aspose
// Developer Intent: Copy a selected cell range from a source workbook into a new workbook and replace every formula with its evaluated value so the destination contains only static data.
// Use Cases: Generate a read‑only report by copying calculated cells and stripping formulas before distribution. | Archive financial models as immutable snapshots to prevent accidental recalculation. | Create a lightweight data export for downstream systems that cannot process Excel formulas.
// AI Prompts: Write C# code using Aspose.Cells to copy a specific range from one workbook to another and replace all formulas with their calculated values. | Show how to copy multiple ranges, preserve cell formatting, and then remove formulas in the destination worksheet. | Explain the impact of Worksheet.Cells.RemoveFormulas after using Range.CopyData in Aspose.Cells, including performance considerations.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Creates a source workbook, adds numeric data and a SUM formula, defines the source range A1:A3, creates a destination workbook, copies the range with CopyData, removes all formulas using Worksheet.Cells.RemoveFormulas, and saves the result as a static‑value workbook (CopiedValues.xlsx).
class Program
{
    static void Main()
    {
        try
        {
            // ---------- Create source workbook and add data with a formula ----------
            Workbook srcWorkbook = new Workbook();
            Worksheet srcSheet = srcWorkbook.Worksheets[0];
            srcSheet.Name = "Source";

            // Populate some cells
            srcSheet.Cells["A1"].PutValue(10);
            srcSheet.Cells["A2"].PutValue(20);
            // Cell with a formula that sums A1 and A2
            srcSheet.Cells["A3"].Formula = "=SUM(A1:A2)";

            // Define the source range to be copied
            AsposeRange srcRange = srcSheet.Cells.CreateRange("A1:A3");

            // ---------- Create destination workbook ----------
            Workbook destWorkbook = new Workbook();
            Worksheet destSheet = destWorkbook.Worksheets[0];
            destSheet.Name = "Copy";

            // Define a destination range of the same size
            AsposeRange destRange = destSheet.Cells.CreateRange("B1:B3");

            // ---------- Copy the range (including formulas) ----------
            destRange.CopyData(srcRange);

            // ---------- Remove all formulas, leaving only static values ----------
            destSheet.Cells.RemoveFormulas();

            // ---------- Save the resulting workbook ----------
            string outputPath = "CopiedValues.xlsx";
            destWorkbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
