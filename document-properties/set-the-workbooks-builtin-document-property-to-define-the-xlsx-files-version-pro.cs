using System;
using Aspose.Cells;

namespace AsposeCellsVersionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Set the built‑in Version property (format "00.0000")
            workbook.BuiltInDocumentProperties.Version = "12.0000";

            // Save the workbook to an XLSX file (lifecycle: save)
            workbook.Save("VersionDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}