// Title: C# – Insert a Worksheet After a Specific Sheet with Aspose.Cells for .NET
// Description: Shows how to create a workbook, find a worksheet by its name, and insert a new worksheet immediately after it using Aspose.Cells for .NET, then save the result.
// Keywords: Aspose.Cells | C# insert worksheet | add worksheet after sheet | worksheet index | control sheet order | .NET spreadsheet library | InsertWorksheet method | Workbook manipulation
// Common Searches: Aspose.Cells insert worksheet after existing sheet | C# add new worksheet after specific tab | How to place a worksheet after another in Aspose.Cells | Get worksheet index by name Aspose.Cells .NET | Programmatically reorder sheets Aspose.Cells
// Developer Intent: Add a new worksheet directly after a given worksheet to set the desired sheet sequence in a generated workbook.
// Use Cases: Create a summary sheet that must follow a data sheet in a financial report. | Insert an analysis tab immediately after a user‑selected worksheet during runtime. | Maintain a predefined sheet order when dynamically adding new worksheets to a template.
// AI Prompts: Generate C# code with Aspose.Cells that inserts a worksheet after a sheet named 'SheetX' and copies the formatting from the preceding sheet. | Provide an example that retrieves a worksheet index by name and inserts multiple worksheets sequentially after it using Aspose.Cells for .NET. | Explain how to handle the situation where the target sheet name does not exist when inserting a new worksheet with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, find a worksheet by its name, and insert a new worksheet immediately after it using Aspose.Cells for .NET, then save the result.
    public class InsertWorksheetAfterSpecifiedSheet
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add initial worksheets
                workbook.Worksheets.Add("SheetA");
                workbook.Worksheets.Add("SheetB");
                workbook.Worksheets.Add("SheetC");

                // Name of the sheet after which the new sheet should be placed
                string targetSheetName = "SheetB";

                // Get the index of the target sheet (zero‑based)
                int targetIndex = workbook.Worksheets[targetSheetName].Index;

                // Insert a new worksheet immediately after the target sheet
                Worksheet insertedSheet = workbook.Worksheets.Insert(targetIndex + 1, SheetType.Worksheet, "InsertedAfterB");

                // Optional: add some data to the new sheet to verify it was created
                insertedSheet.Cells["A1"].PutValue($"This sheet was inserted after {targetSheetName}");

                // Save the workbook
                string outputPath = "SheetSequenceDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the console application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
