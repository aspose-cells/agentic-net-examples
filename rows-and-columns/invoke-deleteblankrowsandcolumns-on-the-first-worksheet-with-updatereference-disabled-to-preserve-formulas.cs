// Title: Remove empty rows and columns from the first worksheet in Aspose.Cells for .NET while keeping formula references unchanged
// AI Prompts: Write C# code that uses Aspose.Cells to delete all blank rows and columns in the first worksheet without altering existing formula references. | Show how to configure DeleteOptions with UpdateReference set to false and apply it to DeleteBlankRows and DeleteBlankColumns methods.
// Common Searches: Aspose.Cells delete blank rows without affecting formulas in C# | How to keep formula references when removing empty columns using Aspose.Cells .NET | DeleteOptions UpdateReference false example for worksheet cleanup | C# code to remove empty rows and columns from a workbook while preserving formulas | Aspose.Cells DeleteBlankRows DeleteBlankColumns with UpdateReference disabled
// Tags: delete blank rows Aspose.Cells C# | delete blank columns Aspose.Cells C# | preserve formulas DeleteOptions | UpdateReference false Aspose.Cells | worksheet cleanup Aspose.Cells .NET

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a workbook, adds data with intentional blank rows and columns, configures DeleteOptions with UpdateReference set to false, deletes the empty rows and columns on the first worksheet, and saves the cleaned workbook as DeletedBlankRowsAndColumns.xlsx.
    public class DeleteBlankRowsAndColumnsDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data with blank rows and columns
            worksheet.Cells["A1"].PutValue("Header");
            worksheet.Cells["A2"].PutValue("Data1");
            // Row 3 is blank
            worksheet.Cells["A4"].PutValue("Data2"); // Blank row at 3
            // Column B is blank
            worksheet.Cells["C1"].PutValue("ColC Header");
            worksheet.Cells["C2"].PutValue("ColC Data");

            // Create DeleteOptions with UpdateReference disabled (false) to preserve formulas
            DeleteOptions deleteOptions = new DeleteOptions
            {
                UpdateReference = false
            };

            // Delete blank rows using the options
            worksheet.Cells.DeleteBlankRows(deleteOptions);

            // Delete blank columns using the same options
            worksheet.Cells.DeleteBlankColumns(deleteOptions);

            // Save the modified workbook
            workbook.Save("DeletedBlankRowsAndColumns.xlsx", SaveFormat.Xlsx);
        }
    }
}
