// Title: Copy a worksheet by name with AddCopy and insert it after another sheet – Aspose.Cells for .NET
// Description: Demonstrates how to use Workbook.Worksheets.AddCopy(string sheetName) to duplicate a worksheet, rename the copy, compute the target sheet index, and move the new sheet directly after a specified worksheet before saving the workbook.
// Keywords: Aspose.Cells | AddCopy overload | copy worksheet by name | insert worksheet after target | Worksheet.MoveTo | C# | .NET workbook manipulation | worksheet index
// Common Searches: Aspose.Cells AddCopy by sheet name | how to insert copied worksheet after another sheet | move worksheet to specific position C# Aspose | copy and reorder worksheets Aspose.Cells | Worksheet.MoveTo example .NET
// Developer Intent: Duplicate a worksheet using AddCopy(string) and place the copy immediately after a chosen sheet.
// Use Cases: Create monthly report tabs by copying a template sheet and inserting each copy after the month‑summary tab. | Generate a backup of a data entry worksheet and position it right after the original for side‑by‑side review. | Build scenario analysis workbooks by copying a base model sheet and inserting each scenario sheet after its corresponding control sheet.
// AI Prompts: Write C# code that copies a worksheet with Workbook.Worksheets.AddCopy("SheetName") and moves the copy after a target sheet using Worksheet.MoveTo. | Explain how to recalculate a target worksheet's index after adding a copy and why this is necessary before calling MoveTo. | Show how to copy multiple worksheets by name and insert each copy after different target sheets in a single Aspose.Cells workbook.

using System;
using Aspose.Cells;

namespace WorksheetCopyInsertDemo
{
    // Demonstrates how to use Workbook.Worksheets.AddCopy(string sheetName) to duplicate a worksheet, rename the copy, compute the target sheet index, and move the new sheet directly after a specified worksheet before saving the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a source worksheet named "Source"
            Worksheet sourceSheet = workbook.Worksheets.Add("Source");
            sourceSheet.Cells["A1"].PutValue("Data in source sheet");

            // Add a target worksheet after which the copy will be placed
            Worksheet targetSheet = workbook.Worksheets.Add("Target");
            targetSheet.Cells["A1"].PutValue("Data in target sheet");

            // Copy the "Source" worksheet using the AddCopy overload that takes a sheet name
            int copiedIndex = workbook.Worksheets.AddCopy("Source");
            Worksheet copiedSheet = workbook.Worksheets[copiedIndex];
            copiedSheet.Name = "SourceCopy";

            // Move the copied sheet to be right after the target sheet
            // Target sheet index may have changed after adding the copy, so recalculate
            int targetIndex = workbook.Worksheets["Target"].Index;
            copiedSheet.MoveTo(targetIndex + 1);

            // Save the workbook
            workbook.Save("WorksheetCopyInsertDemo.xlsx");

            Console.WriteLine("Worksheet copied and inserted successfully.");
        }
    }
}
