// Title: Copy rows with images and drawings using Aspose.Cells in C#
// Description: Demonstrates how to load a source workbook, create a destination workbook, calculate the number of rows to copy, and use Worksheet.Cells.CopyRows with the default copy mode to transfer rows while keeping all embedded pictures and drawing objects intact, then save the result.
// Keywords: Aspose.Cells CopyRows C# | copy rows with pictures | preserve drawing objects | worksheet row duplication | Excel image copy .NET | default copy behavior Aspose
// Common Searches: Aspose.Cells copy rows keep images | CopyRows preserve drawings C# | Transfer rows with embedded pictures Aspose | How to duplicate rows with shapes in .NET Excel | Copy rows between workbooks Aspose.Cells
// Developer Intent: Duplicate rows from one worksheet to another without losing any embedded images or drawing objects.
// Use Cases: Generate client‑specific reports by cloning a template sheet that contains logos and decorative shapes. | Move data rows together with their associated charts from a master workbook to a summary file. | Create a filtered version of a marketing sheet while retaining background graphics and watermarks.
// AI Prompts: Write C# code that uses Aspose.Cells to copy a block of rows from one worksheet to another, ensuring all pictures and shapes are retained. | Explain the default behavior of Worksheet.Cells.CopyRows regarding embedded images and drawing objects, and list any options to modify this behavior.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to load a source workbook, create a destination workbook, calculate the number of rows to copy, and use Worksheet.Cells.CopyRows with the default copy mode to transfer rows while keeping all embedded pictures and drawing objects intact, then save the result.
public class CopyRowsWithImagesDemo
{
    public static void Main()
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        const string sourcePath = "source.xlsx";
        const string outputPath = "output.xlsx";

        // Ensure the source file exists; create a minimal workbook if it does not.
        if (!File.Exists(sourcePath))
        {
            var tempWb = new Workbook();
            var tempSheet = tempWb.Worksheets[0];
            tempSheet.Cells["A1"].PutValue("Sample Data");
            tempWb.Save(sourcePath);
        }

        // Load the source workbook that contains images and drawing objects.
        Workbook sourceWorkbook = new Workbook(sourcePath);
        Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

        // Create a new workbook that will receive the copied rows.
        Workbook destinationWorkbook = new Workbook();
        Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

        // Determine how many rows need to be copied (all rows that have data/formats).
        int rowsToCopy = sourceSheet.Cells.MaxDisplayRange.RowCount;

        // Copy rows from the source sheet to the destination sheet.
        // Default copy behavior includes images and drawings.
        destinationSheet.Cells.CopyRows(sourceSheet.Cells, 0, 0, rowsToCopy);

        // Save the resulting workbook.
        destinationWorkbook.Save(outputPath);
        Console.WriteLine($"Rows copied successfully. Output saved to '{outputPath}'.");
    }
}
