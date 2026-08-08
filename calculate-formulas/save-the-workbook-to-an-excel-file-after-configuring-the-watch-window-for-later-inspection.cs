// Title: Create and Save an Aspose.Cells Workbook with a SUM Formula in C# (Watch Window Unavailable)
// Description: Demonstrates how to instantiate a new Aspose.Cells Workbook in C#, write numeric values to cells A1 and A2, assign a SUM formula to A3, handle potential exceptions, and save the file as "WatchWindowDemo.xlsx" in XLSX format. The example also notes that the current Aspose.Cells library does not expose a Watch Window API.
// Keywords: Aspose.Cells C# save workbook | Aspose.Cells add formula | Aspose.Cells watch window missing | Aspose.Cells create workbook .NET | Aspose.Cells exception handling | Aspose.Cells export to XLSX
// Common Searches: How to save a workbook as .xlsx using Aspose.Cells for .NET | Aspose.Cells set SUM formula in C# | Does Aspose.Cells provide a Watch Window API | Create workbook and add data with Aspose.Cells | Error handling when saving Aspose.Cells workbook
// Developer Intent: Generate a workbook, insert numeric values and a SUM formula, and persist the result as an Excel (.xlsx) file while handling any runtime errors.
// Use Cases: Automate the production of a simple financial sheet that calculates totals before distribution. | Create a template workbook with pre‑filled data and formulas for downstream processing in other systems. | Save a spreadsheet for later inspection or debugging when a Watch Window feature is not yet available.
// AI Prompts: Write C# code with Aspose.Cells to create a workbook, add values to A1 and A2, set a SUM formula in A3, and save it as an .xlsx file with try‑catch error handling. | Explain alternative ways to debug cell values and formulas in Aspose.Cells when the Watch Window API is absent. | Show how to programmatically check Aspose.Cells release notes for a Watch Window feature and suggest fallback logging techniques.

using System;
using Aspose.Cells;

// Demonstrates how to instantiate a new Aspose.Cells Workbook in C#, write numeric values to cells A1 and A2, assign a SUM formula to A3, handle potential exceptions, and save the file as "WatchWindowDemo.xlsx" in XLSX format. The example also notes that the current Aspose.Cells library does not expose a Watch Window API.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet and add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].Formula = "=SUM(A1:A2)";

            // Aspose.Cells does not expose a WatchWindow API in current versions.
            // The watch window functionality is therefore omitted.

            // Save the workbook to an Excel file (lifecycle rule: save)
            string outputFile = "WatchWindowDemo.xlsx";
            workbook.Save(outputFile, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
