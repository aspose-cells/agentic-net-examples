// Title: C# – Delete Blank Columns on First Worksheet with DeleteOptions.UpdateReference in Aspose.Cells
// Description: Demonstrates how to create a workbook, leave column B empty, and remove all blank columns from the first worksheet using Cells.DeleteBlankColumns with DeleteOptions.UpdateReference set to true. The example shows the column count before and after deletion, updates any cell references, and saves the file as an XLSX document.
// Keywords: Aspose.Cells C# delete blank columns | DeleteOptions UpdateReference | Cells.DeleteBlankColumns example | remove empty columns Aspose.Cells .NET | first worksheet column deletion | update formulas after column removal | Aspose.Cells workbook cleanup
// Common Searches: Aspose.Cells delete empty columns with reference update | C# DeleteBlankColumns DeleteOptions.UpdateReference sample | how to remove blank columns from first worksheet Aspose.Cells | update formulas when deleting columns Aspose.Cells .NET | Aspose.Cells column cleanup code example
// Developer Intent: Remove every empty column from the first worksheet while automatically adjusting all cell references.
// Use Cases: Strip placeholder columns from data‑driven workbooks before publishing. | Refresh a template by deleting unused columns and keeping formulas accurate. | Reduce file size and improve readability in automated report generation.
// AI Prompts: Generate C# code that deletes blank columns on the first worksheet using Aspose.Cells with DeleteOptions.UpdateReference enabled. | Explain how DeleteOptions.UpdateReference affects formulas when blank columns are removed in Aspose.Cells. | Show how to log MaxDataColumn before and after calling Cells.DeleteBlankColumns in a .NET example.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, leave column B empty, and remove all blank columns from the first worksheet using Cells.DeleteBlankColumns with DeleteOptions.UpdateReference set to true. The example shows the column count before and after deletion, updates any cell references, and saves the file as an XLSX document.
    public class DeleteBlankColumnsWithUpdateReferenceDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate data with a blank column (column B will be blank)
                cells["A1"].PutValue("Column A");
                cells["C1"].PutValue("Column C"); // Column B is intentionally left blank
                cells["A2"].PutValue(1);
                cells["C2"].PutValue(3);

                // Display column count before deletion
                Console.WriteLine($"Before deletion, MaxDataColumn: {cells.MaxDataColumn}"); // Expected 2 (0‑based)

                // Set DeleteOptions with UpdateReference = true
                DeleteOptions options = new DeleteOptions
                {
                    UpdateReference = true
                };

                // Delete all blank columns using the specified options
                cells.DeleteBlankColumns(options);

                // Display column count after deletion
                Console.WriteLine($"After deletion, MaxDataColumn: {cells.MaxDataColumn}"); // Expected 1 (0‑based)

                // Prepare output path
                string outputFile = "DeleteBlankColumns_Output.xlsx";
                string outputPath = Path.GetFullPath(outputFile);

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
