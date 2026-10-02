// Title: Delete rows and automatically update formula references with DeleteOptions in Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a DeleteOptions object with UpdateReference set to true and calls Worksheet.Cells.DeleteRows to remove specific rows while keeping dependent formulas correct. | Demonstrate how to configure DeleteOptions to choose whether formulas are updated or left as #REF! when deleting rows in an Aspose.Cells workbook.
// Common Searches: Aspose.Cells C# delete rows and keep SUM formula accurate | How to use DeleteOptions.UpdateReference when removing rows in a .NET workbook | C# example for deleting rows without breaking cell references in Aspose.Cells | Control formula reference updates during row deletion with Aspose.Cells DeleteRows | Delete first two rows and update formulas using Aspose.Cells DeleteOptions
// Tags: Aspose.Cells DeleteRows reference update | C# DeleteOptions formula handling | row deletion preserving formulas Aspose.Cells | UpdateReference option example | Aspose.Cells workbook row removal

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a DeleteOptions instance with UpdateReference enabled, builds a workbook with sample values and a SUM formula, deletes the first two rows using Worksheet.Cells.DeleteRows while updating formula references, and saves the workbook as DeletedRows.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a DeleteOptions instance to control reference updating
            DeleteOptions deleteOptions = new DeleteOptions
            {
                // Set UpdateReference to true to update formulas that refer to the deleted cells
                // Set to false if you want the formulas to remain unchanged (they will become #REF!)
                UpdateReference = true
            };

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            // Set a formula referencing A1 and A2
            sheet.Cells["A3"].Formula = "=SUM(A1:A2)";

            // Delete the first two rows (rows 0 and 1) while updating references
            sheet.Cells.DeleteRows(0, 2, deleteOptions.UpdateReference);

            // Save the workbook
            string outputPath = "DeletedRows.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
