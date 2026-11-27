// Title: Replace cells containing "Draft" with "Final" in a named range using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells replace text in a named range Excel workbook | update cells that equal "Draft" to "Final" using Aspose.Cells .NET | iterate through a named range and change string values with Aspose.Cells | how to modify a publishing range in Excel using Aspose.Cells C# | Aspose.Cells .NET change specific cell values in a defined range
// Tags: named range cell iteration Aspose.Cells | string value replacement Aspose.Cells .NET | publish range cell update Aspose.Cells | Excel workbook text edit Aspose.Cells | C# modify cells by value Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook, accesses the named range "PublishRange", iterates over each cell in that range, and replaces any cell whose string value is exactly "Draft" with "Final" before saving the workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range used for publishing
            Name publishRange = workbook.Worksheets.Names["PublishRange"];
            if (publishRange == null)
            {
                Console.WriteLine("Named range 'PublishRange' not found.");
                return;
            }

            // Get the cell range object represented by the named range
            Aspose.Cells.Range range = publishRange.GetRange();

            // Iterate through each cell in the range and replace "Draft" with "Final"
            foreach (Cell cell in range)
            {
                if (cell.Type == CellValueType.IsString && cell.StringValue == "Draft")
                {
                    cell.PutValue("Final");
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
