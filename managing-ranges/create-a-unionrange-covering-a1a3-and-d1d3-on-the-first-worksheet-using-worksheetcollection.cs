// Title: Create a multi‑area union range (A1:A3 and D1:D3) on the first worksheet using Aspose.Cells in C#
// AI Prompts: Generate C# code that uses Aspose.Cells to define a union range covering A1:A3 and D1:D3 on the first worksheet and writes a value into each cell. | Show how to call WorksheetCollection.CreateRange for a non‑contiguous range and iterate over its cells to set data with Aspose.Cells for .NET.
// Common Searches: asp.net aspose.cells create union range A1:A3 D1:D3 | c# how to define a multi‑area range with WorksheetCollection.CreateRange | iterate over non‑contiguous cells Aspose.Cells .NET example | set same value to cells A1:A3 and D1:D3 using Aspose.Cells C# | save workbook after populating a union range with Aspose.Cells
// Tags: create union range Aspose.Cells | WorksheetCollection CreateRange multi‑area | populate non‑contiguous cells C# | Aspose.Cells write values to range | save workbook as xlsx Aspose.Cells

using Aspose.Cells;
using System;

// The program creates a new Workbook, accesses the first worksheet via the Worksheets collection, defines a union range covering A1:A3 and D1:D3 with CreateRange, writes "Sample" into each cell of that range, and saves the file as UnionRangeOutput.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet via the WorksheetCollection
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a range that covers A1:A3 and D1:D3 (union of two areas)
            // Use fully qualified Aspose.Cells.Range to avoid conflict with System.Range
            Aspose.Cells.Range unionRange = worksheet.Cells.CreateRange("A1:A3,D1:D3");

            // Example operation: set a value in each cell of the range
            foreach (Cell cell in unionRange)
            {
                cell.PutValue("Sample");
            }

            // Save the workbook to a file
            workbook.Save("UnionRangeOutput.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
