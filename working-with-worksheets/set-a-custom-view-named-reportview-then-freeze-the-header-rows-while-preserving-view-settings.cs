// Title: Freeze the first row (header) of an Excel worksheet while keeping existing view settings with Aspose.Cells for .NET
// Common Searches: how to freeze the top row in an Excel file using Aspose.Cells C# | preserve worksheet view settings when applying FreezePanes with Aspose.Cells | Aspose.Cells C# freeze header rows without affecting custom view | freeze panes programmatically in .NET while keeping existing view options | example of using FreezePanes to lock header row in Aspose.Cells workbook
// Tags: FreezePanes method preserve worksheet view Aspose.Cells | freeze header row Excel .NET Aspose.Cells | maintain custom view settings after freezing panes C# | Aspose.Cells workbook freeze first row example | Excel worksheet view preservation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing workbook (or creates a new one), accesses the first worksheet, freezes the first row using the FreezePanes method, and saves the file, ensuring any existing view settings remain unchanged.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook if it exists; otherwise create a new one.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // default workbook with one worksheet
            }

            // Work with the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Freeze the first row (header). Use the overload with four parameters.
            // Parameters: row, column, totalRows, totalColumns.
            sheet.FreezePanes(1, 0, 0, 0);

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
