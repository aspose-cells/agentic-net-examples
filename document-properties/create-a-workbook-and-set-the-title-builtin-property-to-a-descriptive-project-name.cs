using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Set the built‑in Title property to a descriptive project name
        workbook.BuiltInDocumentProperties.Title = "Project Alpha – Design Specification";

        // Output the set title for verification
        Console.WriteLine("Workbook Title: " + workbook.BuiltInDocumentProperties.Title);

        // Save the workbook to a file
        workbook.Save("ProjectAlpha.xlsx", SaveFormat.Xlsx);
    }
}