// Title: Delete blank columns from the first worksheet of a .NET workbook using Aspose.Cells default DeleteBlankColumns behavior
// AI Prompts: Generate C# code that creates a workbook, fills data with a blank column, calls worksheet.Cells.DeleteBlankColumns() without changing formulas, and saves the file. | Show how to eliminate all empty columns from the first sheet in Aspose.Cells while keeping existing cell references unchanged. | Provide a minimal example that demonstrates the default behavior where UpdateReference is false when deleting blank columns.
// Common Searches: Aspose.Cells C# delete blank columns without affecting formulas | How to remove empty columns from the first sheet using DeleteBlankColumns default options | DeleteBlankColumns method example with default DeleteOptions in .NET | remove empty columns from worksheet Aspose.Cells C#
// Tags: Aspose.Cells DeleteBlankColumns method | remove empty worksheet columns .NET | default DeleteOptions false UpdateReference | save workbook as XLSX after column deletion | first worksheet column cleanup Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, adds data to columns A and C leaving column B empty, invokes worksheet.Cells.DeleteBlankColumns() which uses the default DeleteOptions (UpdateReference = false), and saves the result as DeletedBlankColumns.xlsx.
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

                // Populate the worksheet with data that includes blank columns
                // Column A has data, Column B is left blank, Column C has data
                worksheet.Cells["A1"].PutValue("Header A");
                worksheet.Cells["C1"].PutValue("Header C"); // B column is blank
                worksheet.Cells["A2"].PutValue(10);
                worksheet.Cells["C2"].PutValue(30);

                // Delete all blank columns on the worksheet using default options
                // (UpdateReference is false by default)
                worksheet.Cells.DeleteBlankColumns();

                // Save the modified workbook
                workbook.Save("DeletedBlankColumns.xlsx", SaveFormat.Xlsx);
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
