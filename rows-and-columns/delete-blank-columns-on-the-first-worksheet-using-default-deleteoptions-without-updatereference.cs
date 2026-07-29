// Title: Delete blank columns in Aspose.Cells (C#) with default DeleteOptions (no reference update)
// Description: Creates a workbook, adds data to columns A and C leaving column B empty, then calls worksheet.Cells.DeleteBlankColumns() which uses the default DeleteOptions (UpdateReference = false) and saves the result as output.xlsx.
// Keywords: Aspose.Cells delete blank columns | DeleteBlankColumns default options | Aspose.Cells C# remove empty columns | DeleteOptions UpdateReference false | worksheet delete blank columns without shifting formulas
// Common Searches: Aspose.Cells delete empty columns without affecting formulas | DeleteBlankColumns default behavior UpdateReference false | C# remove blank columns from first worksheet Aspose.Cells | How to delete blank columns in Aspose.Cells .NET
// Developer Intent: Remove all empty columns from the first worksheet while keeping existing cell references unchanged.
// Use Cases: Clean up a generated workbook by eliminating unused columns before saving. | Prepare a template by stripping placeholder columns that contain no data. | Automate report creation where some columns may be empty and must be removed without altering formulas.
// AI Prompts: Show C# code to delete blank columns in Aspose.Cells while preserving cell references. | Explain how DeleteOptions.UpdateReference influences DeleteBlankColumns in Aspose.Cells. | Provide an example of using DeleteBlankColumns with default options to clean a worksheet.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds data to columns A and C leaving column B empty, then calls worksheet.Cells.DeleteBlankColumns() which uses the default DeleteOptions (UpdateReference = false) and saves the result as output.xlsx.
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

                // Add sample data with a blank column (column B will be blank)
                worksheet.Cells["A1"].PutValue("Column A");
                worksheet.Cells["C1"].PutValue("Column C");
                worksheet.Cells["A2"].PutValue("Data A");
                worksheet.Cells["C2"].PutValue("Data C");

                // Delete all blank columns using the default DeleteBlankColumns method
                // This uses default DeleteOptions where UpdateReference is false
                worksheet.Cells.DeleteBlankColumns();

                // Save the workbook
                workbook.Save("output.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            DeleteBlankColumnsDemo.Run();
        }
    }
}
