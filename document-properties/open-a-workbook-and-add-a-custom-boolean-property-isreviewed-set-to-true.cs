// Title: Add a custom Boolean document property named IsReviewed to an Excel workbook with Aspose.Cells for .NET
// Common Searches: how to add a custom Boolean document property to an Excel file using Aspose.Cells in C# | Aspose.Cells set custom property IsReviewed to true in a workbook | C# code to add and save a custom document property in an .xlsx file with Aspose.Cells | adding custom metadata to an Excel workbook using Aspose.Cells .NET API | save Excel workbook after adding a custom Boolean property with Aspose.Cells
// Tags: custom boolean property Aspose.Cells | set IsReviewed Excel workbook C# | save workbook with custom metadata Aspose.Cells | Excel custom document properties .NET | Aspose.Cells workbook property API

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel file (input.xlsx), adds a custom Boolean document property called "IsReviewed" with a value of true, and saves the updated workbook to output.xlsx. It includes checks for missing files, handles password‑protected workbooks, and catches general exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook (no password supplied). If the file is password‑protected,
            // a CellsException will be thrown and handled in the catch block below.
            var loadOptions = new LoadOptions(); // default options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Add a custom Boolean property named "IsReviewed" and set it to true
            workbook.CustomDocumentProperties.Add("IsReviewed", true);

            // Save the workbook with the new property
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (CellsException ex)
        {
            // Handles errors related to loading/saving the workbook, including password issues
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handles any other unexpected errors
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
