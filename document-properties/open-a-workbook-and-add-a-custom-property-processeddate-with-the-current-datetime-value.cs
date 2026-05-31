using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Create a new workbook (or you could load an existing one)
        Workbook workbook = new Workbook();

        // Add a custom document property named "ProcessedDate" with the current date and time
        workbook.CustomDocumentProperties.Add("ProcessedDate", DateTime.Now);

        // Save the workbook to a file
        workbook.Save("ProcessedWorkbook.xlsx");
    }
}