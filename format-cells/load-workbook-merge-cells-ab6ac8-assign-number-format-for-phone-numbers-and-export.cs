// Title: C# example: Merge cells AB6:AC8 and set custom phone number format '(###) ###‑####' using Aspose.Cells
// Common Searches: Aspose.Cells C# merge cells AB6 AC8 and apply custom phone number format | How to set '(###) ###-####' number format for a merged range in Aspose.Cells .NET | Create style with custom phone number pattern and apply to merged cells using Aspose.Cells for .NET | Load existing workbook, merge AB6:AC8, and format as phone number with Aspose.Cells C# | Apply custom number format to merged cells in an Excel file using Aspose.Cells API
// Tags: merge cells AB6 AC8 Aspose.Cells .NET | phone number pattern style Aspose.Cells | apply style to merged range Aspose.Cells C# | define custom style Aspose.Cells | load workbook and export Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The program loads an existing Excel workbook (or creates a new one), merges the range AB6:AC8 on the first worksheet, defines a style with the custom phone number format '(###) ###‑####', applies this style to the merged cells, and saves the workbook to the specified output file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                try
                {
                    workbook = new Workbook(inputPath);
                }
                catch (CellsException)
                {
                    // Fallback for password‑protected files (using empty password).
                    var loadOptions = new LoadOptions(LoadFormat.Xlsx) { Password = string.Empty };
                    workbook = new Workbook(inputPath, loadOptions);
                }
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one worksheet
            }

            // Get the first worksheet (index 0)
            var worksheet = workbook.Worksheets[0];

            // Merge cells AB6:AC8 (rows and columns are zero‑based)
            worksheet.Cells.Merge(5, 27, 3, 2); // startRow, startColumn, totalRows, totalColumns

            // Create a style with the phone number format '(###) ###-####'
            var phoneStyle = workbook.CreateStyle();
            phoneStyle.Custom = "(###) ###-####";

            // Apply the style to the merged range AB6:AC8
            var range = worksheet.Cells.CreateRange("AB6", "AC8");
            var flag = new StyleFlag { NumberFormat = true };
            range.ApplyStyle(phoneStyle, flag);

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
