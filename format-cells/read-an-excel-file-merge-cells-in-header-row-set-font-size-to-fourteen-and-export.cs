// Title: C# example: Merge header cells A1:D1, set 14‑point font, and save workbook with Aspose.Cells
// Common Searches: how to merge cells A1 to D1 and apply a 14 point font using Aspose.Cells in C# | Aspose.Cells C# code to style a merged header row and export to a new Excel file | load an existing workbook, merge first‑row cells, set header font size, and save with Aspose.Cells .NET | C# Aspose.Cells example for merging header cells and setting font size before saving | handling password‑protected Excel load with Aspose.Cells while formatting header row
// Tags: merge range A1:D1 Aspose.Cells C# | set header font size 14 Aspose.Cells | load workbook with empty password fallback Aspose.Cells | save workbook to XLSX Aspose.Cells .NET | apply style to merged cells Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel file (or creates a new workbook if missing), merges cells A1:D1 on the first worksheet, creates a style with a 14‑point font, applies the style to the merged header cell, and saves the workbook as output.xlsx, including handling of password‑protected files.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        Workbook workbook;

        // Load existing workbook if the file exists; otherwise create a new one
        if (File.Exists(inputPath))
        {
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception)
            {
                // If loading fails (e.g., password‑protected), try with an empty password
                var loadOptions = new LoadOptions(LoadFormat.Xlsx) { Password = string.Empty };
                try
                {
                    workbook = new Workbook(inputPath, loadOptions);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load workbook: {ex.Message}");
                    return;
                }
            }
        }
        else
        {
            // No source file – start with a fresh workbook
            workbook = new Workbook();
        }

        try
        {
            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells A1:D1 (row 0, column 0, 1 row, 4 columns)
            sheet.Cells.Merge(0, 0, 1, 4);

            // Create a style with font size 14
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.Size = 14;

            // Apply the style to the merged cell (top‑left cell of the range)
            StyleFlag styleFlag = new StyleFlag { All = true };
            sheet.Cells[0, 0].SetStyle(headerStyle, styleFlag);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
        }
    }
}
