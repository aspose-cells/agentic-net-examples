// Title: Delete rows in an Excel worksheet while preserving formulas by ensuring UpdateReference is false for each DeleteRows call using Aspose.Cells for .NET
// AI Prompts: Generate C# code that deletes the second and third rows of the first worksheet with Aspose.Cells, passing false for the updateReference parameter on both DeleteRows calls to keep all formulas unchanged. | Demonstrate how to verify that DeleteOptions.UpdateReference remains false before performing a second row deletion to avoid shifting cell references in an Aspose.Cells workbook.
// Common Searches: Aspose.Cells DeleteRows keep formulas unchanged | C# delete multiple rows without updating cell references in Excel | How to prevent formula reference changes when removing rows with Aspose.Cells | Check UpdateReference flag before second DeleteRows call in Aspose.Cells
// Tags: delete rows without updating references Aspose.Cells | DeleteRows updateReference false C# | preserve formulas during row deletion Aspose.Cells | Aspose.Cells DeleteOptions.UpdateReference usage | sequential row removal Excel workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file, deletes the second row and then the third row on the first worksheet using DeleteRows with the updateReference argument set to false for both operations, ensuring that all formulas and cell references remain unchanged, and saves the result to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // First deletion (delete the second row) without updating references
            workbook.Worksheets[0].Cells.DeleteRows(1, 1, false);

            // Second deletion (delete the third row) also without updating references
            workbook.Worksheets[0].Cells.DeleteRows(2, 1, false);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
