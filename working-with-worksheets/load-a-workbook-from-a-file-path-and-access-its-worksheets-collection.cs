// Title: Load an Excel workbook from a file path and enumerate worksheet names with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells load workbook from a specific file path and list worksheet names | How to check if an Excel file exists before opening it with Aspose.Cells in .NET | Iterate through Worksheets collection after loading a workbook using Aspose.Cells C# | Sample code for reading an .xlsx file and printing sheet count with Aspose.Cells | Exception handling when opening Excel files with Aspose.Cells in C#
// Tags: load workbook from file path Aspose.Cells | enumerate worksheets collection Aspose.Cells | verify Excel file existence before loading Aspose.Cells | catch workbook loading exceptions Aspose.Cells | display worksheet count and names C# Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks whether the specified Excel file exists, loads it with Aspose.Cells' Workbook class, accesses the Worksheets collection, prints the total number of worksheets and each sheet's name, and demonstrates proper exception handling for file‑loading errors.
class Program
{
    static void Main(string[] args)
    {
        // Determine the workbook file path.
        // If a path is supplied as a command‑line argument it is used,
        // otherwise a default file name in the current directory is assumed.
        string filePath = args.Length > 0 ? args[0] : "sample.xlsx";

        // Verify that the file exists before attempting to load it.
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file \"{filePath}\" does not exist.");
            return;
        }

        try
        {
            // Load the workbook from the specified file.
            using (Workbook workbook = new Workbook(filePath))
            {
                // Access the collection of worksheets.
                WorksheetCollection worksheets = workbook.Worksheets;

                Console.WriteLine($"Workbook \"{Path.GetFileName(filePath)}\" loaded successfully.");
                Console.WriteLine($"Total worksheets: {worksheets.Count}");

                // Iterate through each worksheet and display its name.
                for (int i = 0; i < worksheets.Count; i++)
                {
                    Worksheet sheet = worksheets[i];
                    Console.WriteLine($"Worksheet {i + 1}: \"{sheet.Name}\"");
                }
            }
        }
        catch (Exception ex)
        {
            // Report any exception that occurs during loading or processing.
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
        }
    }
}
