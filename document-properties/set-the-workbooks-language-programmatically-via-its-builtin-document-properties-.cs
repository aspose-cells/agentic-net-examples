using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class SetWorkbookLanguage
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the language built‑in document property
        workbook.BuiltInDocumentProperties.Language = "en-US";

        // Example of setting another property (optional)
        workbook.BuiltInDocumentProperties.Author = "John Doe";

        // Save the workbook to an XLSX file
        workbook.Save("WorkbookWithLanguage.xlsx");
    }
}