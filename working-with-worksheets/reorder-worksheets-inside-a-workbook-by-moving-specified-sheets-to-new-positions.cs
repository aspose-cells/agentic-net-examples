using System;
using Aspose.Cells;

class ReorderWorksheetsDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample worksheets
        workbook.Worksheets.Add("First");
        workbook.Worksheets.Add("Second");
        workbook.Worksheets.Add("Third");
        workbook.Worksheets.Add("Fourth");

        // Put a marker in each sheet to identify them after reordering
        workbook.Worksheets["First"].Cells["A1"].PutValue("First Sheet");
        workbook.Worksheets["Second"].Cells["A1"].PutValue("Second Sheet");
        workbook.Worksheets["Third"].Cells["A1"].PutValue("Third Sheet");
        workbook.Worksheets["Fourth"].Cells["A1"].PutValue("Fourth Sheet");

        // Move "Third" sheet to the first position (index 0)
        Worksheet thirdSheet = workbook.Worksheets["Third"];
        thirdSheet.MoveTo(0);

        // Move "Fourth" sheet to index 2 (after "First" and "Second")
        Worksheet fourthSheet = workbook.Worksheets["Fourth"];
        fourthSheet.MoveTo(2);

        // Save the reordered workbook
        workbook.Save("ReorderedWorkbook.xlsx");
    }
}