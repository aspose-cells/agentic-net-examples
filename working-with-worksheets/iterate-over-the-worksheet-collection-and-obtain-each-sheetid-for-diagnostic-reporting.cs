using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (empty workbook)
        Workbook workbook = new Workbook();

        // Add sample worksheets for demonstration purposes
        workbook.Worksheets[0].Name = "FirstSheet";
        workbook.Worksheets.Add("SecondSheet");
        workbook.Worksheets.Add("ThirdSheet");

        // Get the worksheet collection
        WorksheetCollection worksheets = workbook.Worksheets;

        // Iterate over each worksheet and retrieve its internal SheetId (TabId)
        for (int i = 0; i < worksheets.Count; i++)
        {
            Worksheet sheet = worksheets[i];
            int sheetId = sheet.TabId; // internal identifier for the sheet
            Console.WriteLine($"Worksheet Name: {sheet.Name}, Index: {sheet.Index}, SheetId (TabId): {sheetId}");
        }

        // Save the workbook (unchanged except for added sheets)
        workbook.Save("DiagnosticReport.xlsx");
    }
}