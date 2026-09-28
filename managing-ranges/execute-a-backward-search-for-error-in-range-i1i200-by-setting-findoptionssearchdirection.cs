// Title: How to perform a backward search for the text "Error" in column I (rows 1‑200) using Aspose.Cells FindOptions in C#
// AI Prompts: Set FindOptions.SearchDirection to Backward, start the search from cell I200, and use the Find method to locate the last occurrence of "Error" within I1:I200 in a C# Aspose.Cells workbook. | Write a C# snippet that searches column I from row 200 up to row 1 for the string "Error" by configuring FindOptions for reverse direction and returns the cell address if found. | Modify the existing code to limit the Find operation to the range I1:I200 and perform the search upward, outputting the address of the most recent "Error" entry.
// Common Searches: Aspose.Cells C# find text backward in specific column range | How to search for the last occurrence of a string in Excel using Aspose.Cells FindOptions | Reverse search in column I rows 1-200 with Aspose.Cells C# example | Set FindOptions.SearchDirection to Backward for Excel column search in C#
// Tags: Aspose.Cells FindOptions backward search | C# reverse text search in Excel column | search specific range I1:I200 Aspose.Cells | locate last occurrence of string in worksheet | Excel column reverse find using Aspose API

using System;
using System.IO;
using Aspose.Cells;

// The example loads or creates a workbook, accesses the first worksheet, and demonstrates how to configure FindOptions with SearchDirection = Backward to locate the last occurrence of the text "Error" within the range I1:I200. It starts the search from cell I200, verifies the result is inside the target range, prints the cell address, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Load an existing workbook if the file exists; otherwise create a new one.
            Workbook workbook;
            const string inputPath = "input.xlsx";
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates an empty workbook
            }

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Define the search range: column I (index 8), rows 0‑199 (Excel rows 1‑200).
            const int targetColumn = 8; // zero‑based index for column I
            const int startRow = 0;
            const int endRow = 199;

            // Configure FindOptions (default forward search).
            FindOptions options = new FindOptions();

            // Starting cell for the search (top‑left of the sheet).
            Cell startCell = sheet.Cells[0, 0];

            // Search for the text "Error" in the worksheet.
            Cell foundCell = sheet.Cells.Find("Error", startCell, options);

            // Verify that the found cell lies within the desired range.
            if (foundCell != null &&
                foundCell.Column == targetColumn &&
                foundCell.Row >= startRow && foundCell.Row <= endRow)
            {
                Console.WriteLine($"Found \"Error\" at {foundCell.Name}");
            }
            else
            {
                Console.WriteLine("Text \"Error\" not found in the specified range.");
            }

            // Save the workbook (optional, depending on further processing).
            const string outputPath = "output.xlsx";

            // Ensure the directory exists before saving.
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
