using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Create a new workbook (creation rule)
        Workbook workbook = new Workbook();

        // Add sample data to demonstrate content in the PDF
        workbook.Worksheets[0].Cells["A1"].PutValue("Workbook with protected VBA project");

        // Protect the VBA project, lock it for viewing and set a password
        // This uses the VbaProject.Protect method as documented
        workbook.VbaProject.Protect(true, "MyVbaPassword");

        // Save the workbook as a PDF file (save rule)
        workbook.Save("ProtectedVbaProject.pdf", SaveFormat.Pdf);
    }
}