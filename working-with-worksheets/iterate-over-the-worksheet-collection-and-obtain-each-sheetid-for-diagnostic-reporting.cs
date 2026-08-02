// Title: Get each worksheet’s internal SheetId (TabId) with Aspose.Cells for .NET
// Description: Creates a workbook, adds sample worksheets, then loops through Workbook.Worksheets, reads each sheet’s TabId, name, and index, writes the data to the console, and saves the workbook as WorksheetIdsReport.xlsx.
// Keywords: Aspose.Cells | .NET | C# | Worksheet TabId | SheetId | iterate worksheets | Workbook.Worksheets loop | diagnostic sheet identifier | debug workbook structure | internal sheet ID
// Common Searches: Aspose.Cells get worksheet TabId .NET | How to read internal SheetId of worksheets in Aspose.Cells | Iterate over worksheets and print TabId using C# | Diagnostic worksheet IDs Aspose.Cells | Retrieve worksheet index and TabId for debugging
// Developer Intent: The developer needs to enumerate all worksheets in a workbook and obtain each worksheet’s internal SheetId (TabId) for diagnostic or logging purposes.
// Use Cases: Log worksheet identifiers to confirm correct ordering after programmatic changes. | Create a diagnostic report of worksheet names, indexes, and TabId values for workbook integrity checks. | Validate that newly added sheets receive expected TabId values during workbook generation.
// AI Prompts: Generate C# code that iterates through an Aspose.Cells workbook and stores each worksheet’s TabId, name, and index in a dictionary. | Show how to export worksheet names, indexes, and TabId values to a CSV file using Aspose.Cells. | Explain how to compare saved TabId values with current worksheet TabId to detect modifications after editing a workbook.

using System;
using Aspose.Cells;

namespace WorksheetIdDiagnostic
{
    // Creates a workbook, adds sample worksheets, then loops through Workbook.Worksheets, reads each sheet’s TabId, name, and index, writes the data to the console, and saves the workbook as WorksheetIdsReport.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Add sample worksheets for demonstration
            workbook.Worksheets.Add("Sales");
            workbook.Worksheets.Add("Inventory");
            workbook.Worksheets.Add("Summary");

            // Iterate over the worksheet collection
            WorksheetCollection sheets = workbook.Worksheets;
            for (int i = 0; i < sheets.Count; i++)
            {
                Worksheet sheet = sheets[i];
                // Obtain the internal SheetId (TabId) for diagnostic reporting
                int sheetId = sheet.TabId;
                Console.WriteLine($"Worksheet Name: {sheet.Name}, Index: {sheet.Index}, SheetId (TabId): {sheetId}");
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save("WorksheetIdsReport.xlsx");
        }
    }
}
