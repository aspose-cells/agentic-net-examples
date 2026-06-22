using Aspose.Cells;

class WorkbookProtectionDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add an extra worksheet for demonstration purposes
        workbook.Worksheets.Add("Data");

        // Rename the first worksheet (allowed before protection)
        workbook.Worksheets[0].Name = "Summary";

        // Protect the workbook structure to block adding new worksheets
        // Structure protection does not affect the rename that was already performed
        workbook.Protect(ProtectionType.Structure, "securePwd");

        // Save the protected workbook
        workbook.Save("ProtectedWorkbook.xlsx");
    }
}