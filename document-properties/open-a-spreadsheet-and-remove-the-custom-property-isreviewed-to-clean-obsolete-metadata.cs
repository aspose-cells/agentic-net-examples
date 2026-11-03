// Title: Delete the custom document property "IsReviewed" from an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# remove custom document property from Excel workbook | how to delete the IsReviewed custom property in an .xlsx file with Aspose.Cells | programmatically clear obsolete custom metadata from Excel using Aspose.Cells .NET | C# code to check and delete a specific custom document property in an Excel file | remove custom property without affecting other workbook data Aspose.Cells
// Tags: Aspose.Cells custom property deletion | Excel .xlsx custom metadata cleanup .NET | C# remove workbook custom document property | manage custom document properties Aspose | delete obsolete Excel property programmatically

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, checks its CustomDocumentProperties collection for a property named "IsReviewed", removes the property if it exists, and saves the modified workbook to a new file while handling missing files, password protection, and other exceptions.
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
            // Load the workbook (no password supplied). If the file is password‑protected,
            // a CellsException will be thrown and handled in the catch block.
            Workbook workbook = new Workbook(inputPath);

            // Access custom document properties collection
            var customProps = workbook.CustomDocumentProperties;

            // Remove the custom property named "IsReviewed" if it exists
            if (customProps.Contains("IsReviewed"))
            {
                customProps.Remove("IsReviewed");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (CellsException ex)
        {
            // Handles errors specific to Aspose.Cells, such as password protection
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handles any other unexpected errors
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
