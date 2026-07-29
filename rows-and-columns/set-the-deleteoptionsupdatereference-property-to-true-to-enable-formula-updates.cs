// Title: C# Example: Delete Columns with UpdateReference = true to auto‑adjust formulas using Aspose.Cells
// Description: Demonstrates creating a workbook, adding formulas that reference a column, configuring DeleteOptions with UpdateReference set to true, deleting the column, and saving the file. The formulas are automatically updated, showing how Aspose.Cells for .NET keeps calculations intact after column removal.
// Keywords: Aspose.Cells DeleteOptions | UpdateReference | C# delete column formulas | Aspose.Cells formula adjustment | DeleteColumns with reference update | Aspose.Cells .NET example | DeleteOptions.UpdateReference true
// Common Searches: Aspose.Cells delete column keep formulas | DeleteOptions.UpdateReference C# | How to update cell references after deleting columns Aspose.Cells | C# Aspose.Cells DeleteColumns example | Auto adjust formulas when removing rows Aspose.Cells
// Developer Intent: Enable automatic formula reference updates when deleting rows or columns by setting DeleteOptions.UpdateReference to true.
// Use Cases: Programmatically remove a column that is referenced by formulas without breaking calculations. | Batch delete multiple columns while preserving dependent formula integrity. | Clean up worksheets by deleting empty or obsolete columns and automatically shifting formula references. | Implement spreadsheet cleanup tools that maintain accurate calculations after structural changes.
// AI Prompts: Generate C# code that deletes a column using Aspose.Cells and updates all dependent formulas. | Show how DeleteOptions.UpdateReference affects formulas when removing rows in Aspose.Cells for .NET. | Explain step‑by‑step how to configure DeleteOptions to keep formulas correct after column deletion. | Provide a complete Aspose.Cells example that demonstrates DeleteColumns with reference updating.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates creating a workbook, adding formulas that reference a column, configuring DeleteOptions with UpdateReference set to true, deleting the column, and saving the file. The formulas are automatically updated, showing how Aspose.Cells for .NET keeps calculations intact after column removal.
    public class DeleteOptionsUpdateReferenceDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created and saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data and formulas that reference column A
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["B1"].PutValue(20);
            sheet.Cells["C1"].Formula = "=A1+B1";

            sheet.Cells["A2"].PutValue(30);
            sheet.Cells["B2"].PutValue(40);
            sheet.Cells["C2"].Formula = "=A2+B2";

            // Create DeleteOptions and enable reference updating
            DeleteOptions deleteOptions = new DeleteOptions
            {
                UpdateReference = true // Enable formula reference updates
            };

            // Delete column A (index 0) using the DeleteOptions.
            // Formulas that referenced column A will be automatically adjusted.
            sheet.Cells.DeleteColumns(0, 1, deleteOptions);

            // Save the modified workbook
            string outputPath = "DeleteOptionsUpdateReferenceDemo.xlsx";
            workbook.Save(outputPath);
        }
    }
}
