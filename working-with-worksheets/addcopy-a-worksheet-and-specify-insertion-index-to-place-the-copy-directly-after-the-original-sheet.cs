// Title: Copy a worksheet and insert it right after the original using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to duplicate a worksheet with Workbook.Worksheets.AddCopy, rename the copy, and place it immediately after the source sheet using Worksheet.MoveTo, then save the workbook.
// Keywords: Aspose.Cells | C# | copy worksheet | AddCopy | Worksheet.MoveTo | insert sheet after original | worksheet index | duplicate Excel sheet | Excel automation .NET | Aspose.Cells example
// Common Searches: Aspose.Cells copy worksheet after original C# | AddCopy method place sheet next to source | Move worksheet to specific index Aspose.Cells | duplicate sheet and set position in .NET | how to insert copied worksheet after original
// Developer Intent: Duplicate an existing worksheet and position the new sheet directly after it in the same workbook.
// Use Cases: Create a template sheet and generate month‑specific copies placed right after the template for organized reporting. | Copy a summary sheet and insert it after each data sheet to keep related information together in automated reports. | Generate multiple worksheets from a formatted base sheet, inserting each copy next to its source before applying custom data.
// AI Prompts: Generate C# code with Aspose.Cells that copies a worksheet named "Data" and inserts the copy immediately after the original sheet. | Show an example using Workbook.Worksheets.AddCopy and Worksheet.MoveTo to duplicate a sheet and set its index to original.Index + 1. | Explain error handling strategies when moving a copied worksheet to a specific position with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Demonstrates how to duplicate a worksheet with Workbook.Worksheets.AddCopy, rename the copy, and place it immediately after the source sheet using Worksheet.MoveTo, then save the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first (default) worksheet and set it up as the original sheet
                Worksheet originalSheet = workbook.Worksheets[0];
                originalSheet.Name = "Original";
                originalSheet.Cells["A1"].PutValue("Data in the original sheet");

                // Copy the original worksheet; AddCopy returns the index of the new sheet
                int copiedIndex = workbook.Worksheets.AddCopy(originalSheet.Name);

                // Retrieve the copied worksheet and give it a distinct name
                Worksheet copiedSheet = workbook.Worksheets[copiedIndex];
                copiedSheet.Name = "CopiedAfterOriginal";

                // Move the copied sheet so that it is placed directly after the original sheet
                copiedSheet.MoveTo(originalSheet.Index + 1);

                // Determine output file path
                string outputFile = "WorksheetCopyAfterOriginal.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputFile));

                // Ensure the output directory exists
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to a file
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
