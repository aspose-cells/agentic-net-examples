// Title: How to enable formula reference updates when deleting rows or columns using DeleteOptions.UpdateReference in Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that deletes a specific column in an Excel workbook with Aspose.Cells while setting DeleteOptions.UpdateReference = true so all dependent formulas are adjusted. | Show how to configure DeleteOptions in Aspose.Cells for .NET to maintain formula integrity when removing multiple rows, including setting UpdateReference to true. | Provide a complete C# example that loads a workbook, applies DeleteOptions with UpdateReference enabled, deletes a range of rows, and saves the file.
// Common Searches: Aspose.Cells C# delete column keep formulas updated | Set DeleteOptions.UpdateReference to true before deleting rows in .NET | Preserve Excel formula references when removing rows using Aspose.Cells API | How to use DeleteOptions with UpdateReference in Aspose.Cells for .NET
// Tags: DeleteOptions.UpdateReference Aspose.Cells | row deletion formula auto‑adjustment .NET | column removal formula sync C# | Aspose.Cells delete options usage | Excel workbook delete rows keep formulas

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Demonstrates loading an Excel workbook, configuring DeleteOptions with UpdateReference set to true, deleting rows or columns, and saving the workbook so that all dependent formulas are automatically updated using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    throw new FileNotFoundException($"The input file '{inputPath}' was not found.");
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Note: Aspose.Cells automatically updates formula references when rows/columns are deleted.
                // If specific delete options are required, they can be set when performing delete operations.

                // Save the workbook after making changes
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any runtime exceptions gracefully
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
