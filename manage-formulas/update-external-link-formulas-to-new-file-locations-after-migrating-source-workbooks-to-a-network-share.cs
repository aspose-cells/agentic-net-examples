// Title: Replace old external workbook paths with a new UNC network share in Excel files using Aspose.Cells for .NET (C#)
// Common Searches: how to update external workbook references to a UNC path with Aspose.Cells C# | replace old file share path in Excel formulas using Aspose.Cells .NET | change external link base folder in a workbook programmatically with Aspose.Cells | migrate external link formulas to a new network share in C# | Aspose.Cells update formulas that contain old folder path
// Tags: update external link formulas Aspose.Cells | replace file share path in Excel workbook C# | modify cell formulas programmatically .NET | UNC path conversion for external references Aspose.Cells | batch update workbook formulas after migration

using Aspose.Cells;
using System;
using System.IO;

namespace Example
{
    // The example loads a target workbook, checks its existence, defines the old and new base directories, iterates through all worksheets and cells, replaces any formula containing the old share path with the new UNC network share path, and saves the updated workbook to a new file while handling potential errors.
    class UpdateExternalLinks
    {
        static void Main()
        {
            try
            {
                // Path to the workbook that contains external link formulas
                string sourcePath = @"C:\Workbooks\Target.xlsx";

                // Verify that the source file exists before loading
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(sourcePath);

                // Old and new base paths for the source workbooks
                string oldBasePath = @"C:\OldShare\";
                string newBasePath = @"\\NetworkShare\NewFolder\";

                // Iterate through all worksheets and cells to update formulas
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;
                    foreach (Cell cell in cells)
                    {
                        if (!string.IsNullOrEmpty(cell.Formula) && cell.Formula.Contains(oldBasePath))
                        {
                            cell.Formula = cell.Formula.Replace(oldBasePath, newBasePath);
                        }
                    }
                }

                // Save the updated workbook
                string outputPath = @"C:\Workbooks\Target_Updated.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
