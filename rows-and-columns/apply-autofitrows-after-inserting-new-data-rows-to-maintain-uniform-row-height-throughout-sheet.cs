// Title: C# – AutoFitRows after InsertRows in Aspose.Cells for .NET to keep uniform row heights
// Description: Shows how to create a workbook, add initial data, insert rows at a given index, fill the new rows, and then invoke Worksheet.AutoFitRows() (and AutoFitColumns()) so that all rows and columns automatically resize before the file is saved.
// Keywords: Aspose.Cells | AutoFitRows | InsertRows | C# example | uniform row height | auto fit columns | dynamic row insertion | worksheet.AutoFitRows() | Aspose.Cells for .NET | adjust row height programmatically
// Common Searches: Aspose.Cells AutoFitRows after inserting rows | C# AutoFitRows not applied to new rows | How to auto size rows in Aspose.Cells .NET | AutoFitRows include newly inserted rows | Aspose.Cells set row height automatically
// Developer Intent: Automatically resize row heights after adding new rows so the worksheet maintains a consistent appearance without manual formatting.
// Use Cases: Insert a variable number of rows into a report template, populate them, and call AutoFitRows to ensure the new rows match existing row heights. | Generate a data‑driven spreadsheet where rows are added on‑the‑fly, then apply AutoFitRows (and AutoFitColumns) to improve readability. | Create a reusable workbook, programmatically add data rows, and use AutoFitRows to keep formatting uniform across all sheets.
// AI Prompts: Provide a C# snippet that inserts rows into an Aspose.Cells worksheet, fills the cells, and calls Worksheet.AutoFitRows() to auto‑adjust row heights. | Explain why Worksheet.AutoFitRows() must be called after InsertRows and describe any additional steps needed to include newly added rows. | Generate a step‑by‑step guide for using AutoFitRows together with AutoFitColumns after dynamic row insertion in Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, add initial data, insert rows at a given index, fill the new rows, and then invoke Worksheet.AutoFitRows() (and AutoFitColumns()) so that all rows and columns automatically resize before the file is saved.
    public class AutoFitRowsAfterInsertDemo
    {
        // Entry point for the application
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
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Add some initial data
            cells["A1"].PutValue("Header");
            cells["B1"].PutValue("Value");
            cells["A2"].PutValue("Row 1");
            cells["B2"].PutValue(10);
            cells["A3"].PutValue("Row 2");
            cells["B3"].PutValue(20);

            // Insert two new rows at index 2 (between existing rows 2 and 3)
            cells.InsertRows(2, 2);

            // Populate the newly inserted rows with data
            cells["A3"].PutValue("Inserted Row 1");
            cells["B3"].PutValue(100);
            cells["A4"].PutValue("Inserted Row 2");
            cells["B4"].PutValue(200);

            // AutoFit rows and columns for better visibility
            worksheet.AutoFitRows();
            worksheet.AutoFitColumns();

            // Save the workbook
            string outputPath = "AutoFitRowsAfterInsertDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
    }
}
