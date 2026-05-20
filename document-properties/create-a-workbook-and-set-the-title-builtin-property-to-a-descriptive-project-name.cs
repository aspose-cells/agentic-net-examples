using System;
using Aspose.Cells;

namespace AsposeCellsTitleDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the Workbook constructor - lifecycle create rule)
            Workbook workbook = new Workbook();

            // Set the built‑in Title property to a descriptive project name
            workbook.BuiltInDocumentProperties.Title = "Project XYZ - Financial Report";

            // Optionally display the title to verify
            Console.WriteLine("Workbook Title: " + workbook.BuiltInDocumentProperties.Title);

            // Save the workbook to a file (uses the Save method - lifecycle save rule)
            string outputPath = "ProjectXYZ_Report.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine("Workbook saved to: " + outputPath);
        }
    }
}