// Title: C# example: Merge header cells A1:D1 and set font size to 14 using Aspose.Cells, then save the workbook
// Common Searches: Aspose.Cells C# merge first row cells and apply font size 14 | How to set font size for a merged header range in Excel with Aspose.Cells | C# code to merge cells A1 to D1 and format header font using Aspose.Cells | Load existing Excel file or create new workbook, then merge header cells with Aspose.Cells | Saving a workbook after merging cells and styling the header in Aspose.Cells .NET
// Tags: merge cells Aspose.Cells C# | apply font size Aspose.Cells style | header row formatting Aspose.Cells | load or create workbook Aspose.Cells | save workbook Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel file (or creates a new workbook if the file is missing), merges cells A1 through D1 on the first worksheet, applies a 14‑point font to the merged header range using a Style and StyleFlag, and saves the modified workbook to output.xlsx, with error handling for load and save operations.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                try
                {
                    workbook = new Workbook(inputPath);
                }
                catch (CellsException ex)
                {
                    // If the file cannot be opened (e.g., password protected), create a new workbook
                    Console.WriteLine($"Failed to load \"{inputPath}\": {ex.Message}. Creating a new workbook instead.");
                    workbook = new Workbook();
                }
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                workbook = new Workbook();
            }

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range to merge (A1:D1)
            CellArea headerArea = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 0,
                EndColumn = 3
            };

            // Merge the defined cells
            sheet.Cells.Merge(
                headerArea.StartRow,
                headerArea.StartColumn,
                headerArea.EndRow - headerArea.StartRow + 1,
                headerArea.EndColumn - headerArea.StartColumn + 1);

            // Apply font size 14 to the merged header range
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.Size = 14;
            StyleFlag styleFlag = new StyleFlag { FontSize = true };

            sheet.Cells.CreateRange(
                headerArea.StartRow,
                headerArea.StartColumn,
                headerArea.EndRow - headerArea.StartRow + 1,
                headerArea.EndColumn - headerArea.StartColumn + 1)
                .ApplyStyle(headerStyle, styleFlag);

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
