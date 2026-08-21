// Title: C# – Delete Blank Columns in the First Worksheet with Aspose.Cells (UpdateReference = true)
// Description: Demonstrates how to create a workbook, add data leaving column B empty, configure DeleteOptions with UpdateReference enabled, call worksheet.Cells.DeleteBlankColumns(options) to remove all empty columns from the first sheet, and save the result as output.xlsx.
// Keywords: Aspose.Cells delete blank columns C# | DeleteOptions UpdateReference true | Remove empty columns Aspose.Cells | worksheet.Cells.DeleteBlankColumns example | C# spreadsheet column cleanup | Aspose.Cells .NET delete empty columns | update references after column deletion | Aspose.Cells API DeleteBlankColumns
// Common Searches: Aspose.Cells delete blank columns with UpdateReference | C# remove empty columns from worksheet Aspose.Cells | How to keep formulas after deleting columns in Aspose.Cells | DeleteBlankColumns options C# | Aspose.Cells delete empty columns first sheet
// Developer Intent: Remove all empty columns from the first worksheet while automatically updating cell references.
// Use Cases: Sanitize imported Excel files by stripping out blank columns before analysis. | Prepare workbooks for publishing where stray empty columns affect layout or printing. | Ensure formulas and named ranges stay accurate after column removal in financial models.
// AI Prompts: Write C# code that deletes blank columns in an Aspose.Cells workbook and updates all references. | Show how to delete empty columns from a specific worksheet while preserving named ranges using Aspose.Cells for .NET. | Explain the effect of DeleteOptions.UpdateReference when removing blank columns with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to create a workbook, add data leaving column B empty, configure DeleteOptions with UpdateReference enabled, call worksheet.Cells.DeleteBlankColumns(options) to remove all empty columns from the first sheet, and save the result as output.xlsx.
public class DeleteBlankColumnsDemo
{
    public static void Run()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate the worksheet with data that includes a blank column (column B)
            worksheet.Cells["A1"].PutValue("Column A");
            worksheet.Cells["C1"].PutValue("Column C"); // Column B is intentionally left blank
            worksheet.Cells["A2"].PutValue(1);
            worksheet.Cells["C2"].PutValue(3);

            // Configure DeleteOptions to update references after deletion
            DeleteOptions options = new DeleteOptions
            {
                UpdateReference = true
            };

            // Delete all blank columns using the configured options
            worksheet.Cells.DeleteBlankColumns(options);

            // Ensure the output directory exists
            string outputPath = "output.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main()
    {
        DeleteBlankColumnsDemo.Run();
    }
}
