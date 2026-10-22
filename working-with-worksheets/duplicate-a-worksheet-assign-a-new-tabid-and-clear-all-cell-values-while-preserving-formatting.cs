// Title: Copy an Excel worksheet, set a unique TabId, and clear its cell contents while keeping formatting with Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to copy a specified worksheet, assign a new TabId to the copy, clear only the cell values, and preserve all formatting. | Create a method that takes an input .xlsx file, duplicates a sheet with a new name, updates its TabId, removes all data from the duplicated sheet without affecting styles, and saves the workbook.
// Common Searches: Aspose.Cells how to copy a sheet and change its TabId in C# | clear worksheet data but retain formatting after copying with Aspose.Cells .NET | set unique TabId for duplicated worksheet using Aspose.Cells | C# Aspose.Cells duplicate worksheet and clear contents example
// Tags: copy worksheet with new TabId Aspose.Cells | clear cell contents preserving formatting .NET | duplicate sheet and reset TabId C# | Aspose.Cells worksheet copy clear values | preserve styles while clearing data Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, copies the first worksheet, gives the copy a new name and a unique TabId, clears all cell values while retaining formatting, and saves the modified workbook using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Choose the worksheet to duplicate (e.g., the first worksheet)
            int sourceIndex = 0;
            Worksheet sourceSheet = workbook.Worksheets[sourceIndex];

            // Duplicate the worksheet with a new name
            string newSheetName = sourceSheet.Name + "_Copy";
            // AddCopy expects a sheet name or index, not a Worksheet object
            int copiedIndex = workbook.Worksheets.AddCopy(sourceSheet.Name);
            Worksheet copiedSheet = workbook.Worksheets[copiedIndex];
            copiedSheet.Name = newSheetName;

            // Assign a new, unique TabId to the duplicated worksheet
            copiedSheet.TabId = workbook.Worksheets.Count; // ensures uniqueness

            // Clear all cell values while preserving formatting
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = copiedSheet.Cells.MaxDataRow,
                EndColumn = copiedSheet.Cells.MaxDataColumn
            };
            copiedSheet.Cells.ClearContents(area);

            // Save the workbook with the changes
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
