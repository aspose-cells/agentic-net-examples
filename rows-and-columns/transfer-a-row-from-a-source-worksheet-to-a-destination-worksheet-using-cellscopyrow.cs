// Title: Copy a specific row from one worksheet to another using Aspose.Cells Cells.CopyRow in C#
// Common Searches: Aspose.Cells copy row from Sheet1 to Sheet2 C# example | How to use Cells.CopyRow to transfer a row between worksheets in .NET | Copy row by index to a different worksheet with Aspose.Cells | Create destination worksheet if missing and copy row Aspose.Cells | Save workbook after copying a row with Aspose.Cells C#
// Tags: Aspose.Cells copy row between worksheets | Cells.CopyRow C# usage | copy row by index Aspose.Cells | transfer Excel row to another sheet .NET | create missing worksheet Aspose.Cells | save workbook after row copy Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel file, ensures a destination worksheet exists, then copies a specified row from the source sheet to a target row in the destination sheet using Cells.CopyRow, and finally saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        const string sourcePath = "source.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the source workbook exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Get the source worksheet (by name or index)
            Worksheet sourceSheet = workbook.Worksheets["Sheet1"] ?? workbook.Worksheets[0];

            // Get or create the destination worksheet
            Worksheet destinationSheet = workbook.Worksheets["Sheet2"];
            if (destinationSheet == null)
            {
                destinationSheet = workbook.Worksheets.Add("Sheet2");
            }

            // Define the row indices (0‑based)
            int sourceRowIndex = 2;        // Row to copy from (e.g., third row)
            int destinationRowIndex = 5;   // Row to copy to (e.g., sixth row)

            // Copy the entire row from the source worksheet to the destination worksheet
            // Note: CopyRow is invoked on the destination Cells collection.
            destinationSheet.Cells.CopyRow(sourceSheet.Cells, sourceRowIndex, destinationRowIndex);

            // Save the workbook with the changes
            workbook.Save(outputPath);
            Console.WriteLine($"Row copied successfully. Output saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
