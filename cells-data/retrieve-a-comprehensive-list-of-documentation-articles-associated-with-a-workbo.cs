using System;
using Aspose.Cells;

class RetrieveDocumentationArticles
{
    static void Main()
    {
        // Path to the existing XLSX workbook
        string workbookPath = "SampleWorkbook.xlsx";

        // Load the workbook
        Workbook workbook = new Workbook(workbookPath);

        // Output the names of the worksheets as a simple documentation list
        Console.WriteLine("Worksheets in the workbook:");
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Console.WriteLine("- " + sheet.Name);
        }
    }
}