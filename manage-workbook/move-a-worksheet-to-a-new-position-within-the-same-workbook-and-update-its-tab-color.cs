// Title: C# – Move a worksheet to a new tab position and change its tab color with Aspose.Cells
// Common Searches: Aspose.Cells C# move worksheet to first tab and set tab color | How to reorder Excel worksheets and change tab color using Aspose.Cells for .NET | C# code to move a sheet to a specific index and apply a tab color with Aspose.Cells | Programmatically change worksheet order and tab appearance in an existing workbook Aspose.Cells | Move Sheet2 to position 0 and set LightBlue tab color Aspose.Cells example
// Tags: move worksheet index Aspose.Cells .NET | set worksheet tab color Aspose.Cells | reorder Excel worksheets programmatically | Aspose.Cells change sheet tab appearance | load and save workbook C# Aspose.Cells | handle missing workbook file Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The example loads an existing Excel workbook, moves the worksheet named "Sheet2" to the first tab (index 0), sets its tab color to LightBlue, and saves the workbook, with error handling for missing files and Aspose.Cells exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the worksheet to be moved (by name)
            Worksheet sheet = workbook.Worksheets["Sheet2"];
            if (sheet == null)
            {
                Console.WriteLine("Error: Worksheet \"Sheet2\" does not exist.");
                return;
            }

            // Move the worksheet to the first tab (index 0)
            sheet.MoveTo(0);

            // Set a new tab color for the moved worksheet
            sheet.TabColor = Color.LightBlue;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (CellsException ex)
        {
            // Handles Aspose.Cells specific errors (e.g., password protection)
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handles any other unexpected errors
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
