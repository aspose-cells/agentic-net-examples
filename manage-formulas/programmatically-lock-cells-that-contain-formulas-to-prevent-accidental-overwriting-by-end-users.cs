// Title: How to lock formula cells and protect worksheets with a password using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells lock cells that contain formulas and protect worksheet with password | programmatically set IsLocked style for formula cells in Excel using Aspose.Cells .NET | apply worksheet protection to all sheets after locking formula cells in Aspose.Cells | iterate through used cells to lock formulas in Aspose.Cells workbook | save workbook after applying password protection with Aspose.Cells C#
// Tags: lock formula cells Aspose.Cells .NET | worksheet protection password Aspose.Cells | set cell IsLocked style based on formula Aspose.Cells | iterate used cells Aspose.Cells C# | save protected workbook Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The C# program loads an Excel file with Aspose.Cells, iterates each worksheet and its used cells, sets the IsLocked style to true for cells containing formulas, protects each sheet with a specified password, ensures the output directory exists, and saves the modified workbook, handling any errors that occur.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string password = "YourPassword";

        try
        {
            // Ensure the input file exists; create a blank workbook if it does not.
            if (!File.Exists(inputPath))
            {
                var tempWb = new Workbook();
                tempWb.Save(inputPath);
            }

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Iterate through all used cells.
                foreach (Cell cell in cells)
                {
                    // Lock cells that contain a formula; unlock others.
                    Style style = cell.GetStyle();
                    style.IsLocked = !string.IsNullOrEmpty(cell.Formula);
                    cell.SetStyle(style);
                }

                // Protect the worksheet (oldPassword parameter is optional, pass empty string).
                sheet.Protect(ProtectionType.All, password, string.Empty);
            }

            // Ensure the output directory exists.
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
