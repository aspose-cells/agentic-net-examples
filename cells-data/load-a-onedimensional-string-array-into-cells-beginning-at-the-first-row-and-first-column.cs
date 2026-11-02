// Title: How to load a one‑dimensional string array into the first row of an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# write string array to first row of worksheet | populate Excel cells horizontally from a string[] with Aspose.Cells | C# Aspose.Cells PutValue loop for one‑dimensional array | save workbook after inserting array into first row using Aspose.Cells | create new workbook and add string array to cells[0, column] in .NET
// Tags: string array to Excel cells Aspose.Cells | horizontal array population Cells.PutValue | initialize workbook first worksheet Aspose.Cells | save workbook as Xlsx Aspose.Cells | populate first row from one‑dimensional array C#

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, accesses the first Worksheet, iterates over a one‑dimensional string array, and writes each element into the first row cells using Cells[row, column].PutValue. It ensures the output directory exists and saves the workbook as an Xlsx file.
class Program
{
    static void Main()
    {
        try
        {
            // One‑dimensional string array to load into the worksheet
            string[] data = new string[] { "Apple", "Banana", "Cherry", "Date" };

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (index 0)
            Worksheet sheet = workbook.Worksheets[0];

            // Import the array horizontally starting at the first row (0) and first column (0)
            for (int i = 0; i < data.Length; i++)
            {
                // PutValue writes the string into the cell at row 0, column i
                sheet.Cells[0, i].PutValue(data[i]);
            }

            // Define output file path
            string outputPath = "output.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
