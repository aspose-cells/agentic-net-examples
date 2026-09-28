// Title: How to set a uniform 15‑point row height for every row in all worksheets of an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an existing .xlsx file with Aspose.Cells, iterates through each worksheet, and sets the Height property of every row to 15 points before saving. | Write a .NET method that guarantees every worksheet—including empty ones—contains at least one row with a height of 15 points using Aspose.Cells. | Create a function that applies a consistent 15‑point row height to all rows in a workbook and returns the updated Workbook object.
// Common Searches: Aspose.Cells C# set row height for all rows in a workbook | How to apply the same row height to every worksheet in an Excel file using .NET | C# code to enforce a minimum row height of 15 points on empty sheets with Aspose.Cells | Batch update row heights across multiple worksheets with Aspose.Cells for .NET
// Tags: set row height Aspose.Cells .NET | uniform row height across worksheets | apply 15 point row height C# | process empty worksheets Aspose.Cells | batch row height modification

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook (or creates a new one), loops through each worksheet, and sets every row's Height property to 15 points, ensuring at least one row is processed even on empty sheets, then saves the modified file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one worksheet
            }

            // Iterate over all worksheets and rows to set a uniform row height (e.g., 15 points)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow; // highest row with data
                // Ensure at least one row is processed even if sheet is empty
                int rowsToProcess = Math.Max(maxRow + 1, 1);

                for (int rowIndex = 0; rowIndex < rowsToProcess; rowIndex++)
                {
                    Row row = cells.Rows[rowIndex];
                    row.Height = 15; // set row height in points
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
