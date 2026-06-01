using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        try
        {
            // Create source workbook and populate a range with sample data
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    sourceSheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Define the source range (A1:C3)
            AsposeRange sourceRange = sourceSheet.Cells.CreateRange(0, 0, 3, 3);

            // Create destination workbook
            Workbook destinationWorkbook = new Workbook();
            Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

            // Define the destination range (starting at A1)
            AsposeRange destinationRange = destinationSheet.Cells.CreateRange(0, 0, 3, 3);

            // Copy the source range to the destination range
            destinationRange.Copy(sourceRange);

            // Protect the destination worksheet with a password for editing
            destinationSheet.Protect(ProtectionType.All, "myPassword", null);

            // Save the resulting workbook
            string outputPath = "CopiedAndProtected.xlsx";
            destinationWorkbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}