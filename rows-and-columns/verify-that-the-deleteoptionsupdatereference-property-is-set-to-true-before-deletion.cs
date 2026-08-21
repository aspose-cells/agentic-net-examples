// Title: Aspose.Cells for .NET – Ensure DeleteOptions.UpdateReference is true before deleting a column (C#)
// Description: C# example that builds a workbook, adds values and formulas referencing column A, verifies the UpdateReference flag, enables it if needed, deletes column A with Cells.DeleteColumns, and confirms that formulas are automatically adjusted.
// Keywords: Aspose.Cells DeleteOptions | UpdateReference flag | C# column deletion | formula reference update | Cells.DeleteColumns | Aspose.Cells .NET example | verify DeleteOptions | preserve formulas after delete
// Common Searches: Aspose.Cells keep formulas after deleting a column | Set UpdateReference flag in C# Aspose.Cells | DeleteColumns with reference update Aspose.Cells | Check DeleteOptions before column removal | Aspose.Cells .NET delete column and adjust formulas
// Developer Intent: Enable the UpdateReference flag so that removing a column automatically adjusts any dependent formulas.
// Use Cases: Programmatically confirm and turn on UpdateReference before invoking Cells.DeleteColumns. | Delete a column while maintaining correct formula calculations across the worksheet. | Log when the flag was initially disabled, set it to true, perform the deletion, and verify the updated formulas.
// AI Prompts: Generate C# code using Aspose.Cells that checks the UpdateReference flag, sets it if false, deletes a column, and prints the resulting formulas. | Show how to delete multiple columns with DeleteOptions while ensuring formula references are updated in Aspose.Cells .NET. | Explain how the UpdateReference setting influences formula recalculation when columns are removed.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // C# example that builds a workbook, adds values and formulas referencing column A, verifies the UpdateReference flag, enables it if needed, deletes column A with Cells.DeleteColumns, and confirms that formulas are automatically adjusted.
    public class VerifyDeleteOptionsUpdateReference
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data and formulas that reference column A
            cells["A1"].PutValue(10);
            cells["B1"].Formula = "=A1*2";
            cells["A2"].PutValue(20);
            cells["B2"].Formula = "=A2*2";

            // Create DeleteOptions instance
            DeleteOptions deleteOptions = new DeleteOptions();

            // Ensure UpdateReference is true before deletion
            if (!deleteOptions.UpdateReference)
            {
                deleteOptions.UpdateReference = true;
                Console.WriteLine("DeleteOptions.UpdateReference was false; set to true.");
            }
            else
            {
                Console.WriteLine("DeleteOptions.UpdateReference is already true.");
            }

            // Delete column A (index 0) using the verified DeleteOptions
            // This will also update the formulas in column B accordingly
            cells.DeleteColumns(0, 1, deleteOptions);

            // Output the formula after deletion to confirm it was updated
            Console.WriteLine("Formula in A1 after column deletion: " + cells["A1"].Formula);

            // Save the workbook
            string outputPath = "VerifyDeleteOptionsUpdateReference_Output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to: " + outputPath);
        }
    }
}
