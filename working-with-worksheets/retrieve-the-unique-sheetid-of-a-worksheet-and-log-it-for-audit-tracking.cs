using System;
using Aspose.Cells;

namespace AsposeCellsAuditDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Generate a GUID and assign it as the worksheet's UniqueId
            // (this mimics retrieving an existing unique ID; if the workbook is loaded,
            // the UniqueId would already be present)
            string guid = "{" + Guid.NewGuid().ToString() + "}";
            sheet.UniqueId = guid;

            // Log the UniqueId for audit tracking
            Console.WriteLine($"Worksheet \"{sheet.Name}\" UniqueId: {sheet.UniqueId}");

            // Save the workbook (lifecycle rule: save)
            string outputPath = "AuditWorkbook.xlsx";
            workbook.Save(outputPath);
        }
    }
}