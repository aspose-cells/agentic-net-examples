// Title: Aspose.Cells C# – Verify source range is empty after MoveTo
// Description: This example creates a workbook, fills A1:B2 with data, moves the range to C3:D4 using Range.MoveTo, then uses Range.IsBlank to confirm the original cells are cleared. The result is printed and the workbook is saved.
// Keywords: Aspose.Cells | C# | .NET | Range.MoveTo | Range.IsBlank | validate moved range | clear source cells | Excel automation | cell range relocation
// Common Searches: Aspose.Cells check if source range is blank after MoveTo | C# verify cells cleared after moving a range | Range.IsBlank after moving range in Aspose.Cells | how to test range relocation in .NET Excel library
// Developer Intent: Confirm that calling MoveTo removes all data from the original range.
// Use Cases: Automated unit test to ensure range relocation does not leave duplicate data. | Data migration within a worksheet where the source area must be emptied. | Report generation that moves calculated blocks to a summary section and validates cleanup.
// AI Prompts: Generate an NUnit test that moves a range with Aspose.Cells and asserts sourceRange.IsBlank() is true. | Provide a C# snippet that moves a range, logs source and destination values, and throws an exception if the source is not empty. | Explain how to use Aspose.Cells Range.IsBlank after MoveTo to confirm successful source clearance.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// This example creates a workbook, fills A1:B2 with data, moves the range to C3:D4 using Range.MoveTo, then uses Range.IsBlank to confirm the original cells are cleared. The result is printed and the workbook is saved.
public class MoveRangeValidationDemo
{
    public static void Run()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate the source range A1:B2 with sample data
            cells["A1"].PutValue("Val1");
            cells["B1"].PutValue("Val2");
            cells["A2"].PutValue("Val3");
            cells["B2"].PutValue("Val4");

            // Create a Range object representing the source area
            AsposeRange sourceRange = cells.CreateRange("A1:B2");

            // Move the range to C3:D4 (zero‑based indices: row 2, column 2)
            sourceRange.MoveTo(2, 2);

            // After moving, the original area should be empty.
            // Create a new Range object for the original location and check if it is blank.
            AsposeRange originalArea = cells.CreateRange("A1:B2");
            bool isBlank = originalArea.IsBlank();

            Console.WriteLine("Source range after move is blank: " + isBlank);

            // Verify that the destination contains the moved data
            Console.WriteLine("Destination C3 value: " + cells["C3"].StringValue);
            Console.WriteLine("Destination D4 value: " + cells["D4"].StringValue);

            // Save the workbook
            string outputPath = "MoveRangeValidationOutput.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to " + outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}

public class Program
{
    public static void Main()
    {
        MoveRangeValidationDemo.Run();
    }
}
