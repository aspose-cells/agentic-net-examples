using System;
using System.IO;
using Aspose.Cells;

class CopyRangeAndProtectWorkbook
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string outputPath = "output.xlsx";

            // Verify source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook
            Workbook sourceWb = new Workbook(sourcePath);

            // Define the source range (A1:C5)
            Worksheet sourceSheet = sourceWb.Worksheets[0];
            Cells sourceCells = sourceSheet.Cells;
            Aspose.Cells.Range sourceRange = sourceCells.CreateRange("A1:C5");

            // Create a new (empty) destination workbook
            Workbook destWb = new Workbook();

            // Define the destination range (A1:C5)
            Worksheet destSheet = destWb.Worksheets[0];
            Cells destCells = destSheet.Cells;
            Aspose.Cells.Range destRange = destCells.CreateRange("A1:C5");

            // Copy the source range into the destination range
            destRange.Copy(sourceRange);

            // Protect the workbook structure with a password
            destWb.Protect(ProtectionType.Structure, "password123");

            // Save the resulting workbook
            destWb.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}