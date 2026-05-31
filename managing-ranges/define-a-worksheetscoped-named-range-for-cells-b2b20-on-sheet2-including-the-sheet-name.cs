using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Ensure a worksheet named "Sheet2" exists
        Worksheet sheet2 = workbook.Worksheets["Sheet2"];
        if (sheet2 == null)
        {
            sheet2 = workbook.Worksheets.Add("Sheet2");
        }

        // Add a new name to the workbook's name collection
        int nameIdx = workbook.Worksheets.Names.Add("MyRange");
        Name namedRange = workbook.Worksheets.Names[nameIdx];

        // Set the scope to Sheet2 (SheetIndex is one‑based, worksheet index is zero‑based)
        namedRange.SheetIndex = sheet2.Index + 1;

        // Define the range reference, including the sheet name
        namedRange.RefersTo = "=Sheet2!$B$2:$B$20";

        // Save the workbook
        workbook.Save("WorksheetScopedNamedRange.xlsx");
    }
}