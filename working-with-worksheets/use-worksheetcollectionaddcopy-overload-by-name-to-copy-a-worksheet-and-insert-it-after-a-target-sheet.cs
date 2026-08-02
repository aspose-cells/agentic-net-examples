// Title: Copy a Worksheet by Name and Insert After Target Sheet – Aspose.Cells for .NET (C#)
// Description: Demonstrates how to use Aspose.Cells for .NET to copy a worksheet using the Worksheets.AddCopy(string) overload, rename the copy, move it to a position directly after a specified target sheet, and save the workbook.
// Keywords: Aspose.Cells | C# | Worksheet AddCopy overload | copy worksheet by name | insert worksheet after another | move worksheet index | save workbook Aspose.Cells | .NET spreadsheet example
// Common Searches: Aspose.Cells copy worksheet by name C# | Insert copied sheet after specific sheet Aspose.Cells | Worksheets.AddCopy example .NET | Move worksheet to specific index Aspose.Cells | How to rename and reposition a copied worksheet in C#
// Developer Intent: Copy an existing worksheet identified by its name and place the duplicate immediately after a chosen target worksheet in the same workbook.
// Use Cases: Create department‑specific reports by duplicating a template sheet and inserting each copy after the department's summary sheet. | Generate a financial workbook where a data sheet is copied and positioned right after a chart sheet to keep related information together. | Automate workbook assembly by copying a configuration sheet and moving it after a header sheet before exporting the file.
// AI Prompts: Write C# code using Aspose.Cells to copy a worksheet named "Template" and insert it directly after a worksheet named "Summary" in an existing workbook. | Show an example that calls Worksheets.AddCopy(string) followed by MoveTo to reorder sheets, then saves the workbook as an XLSX file. | Explain step‑by‑step how to rename a copied worksheet and place it at a specific index using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // Demonstrates how to use Aspose.Cells for .NET to copy a worksheet using the Worksheets.AddCopy(string) overload, rename the copy, move it to a position directly after a specified target sheet, and save the workbook.
    class WorksheetCopyInsertDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add a source worksheet named "Source"
                Worksheet source = workbook.Worksheets.Add("Source");
                source.Cells["A1"].PutValue("Data in source sheet");

                // Add a target worksheet after which the copy will be placed
                Worksheet target = workbook.Worksheets.Add("Target");
                target.Cells["A1"].PutValue("Data in target sheet");

                // Copy the source worksheet using the AddCopy overload that takes a sheet name
                int copiedIndex = workbook.Worksheets.AddCopy("Source");
                Worksheet copied = workbook.Worksheets[copiedIndex];
                copied.Name = "SourceCopy";

                // Move the copied sheet to the position right after the target sheet
                int insertPosition = target.Index + 1;
                copied.MoveTo(insertPosition);

                // Save the workbook
                workbook.Save("WorksheetCopyInsertDemo.xlsx");
                Console.WriteLine("Workbook saved successfully as WorksheetCopyInsertDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            WorksheetCopyInsertDemo.Run();
        }
    }
}
