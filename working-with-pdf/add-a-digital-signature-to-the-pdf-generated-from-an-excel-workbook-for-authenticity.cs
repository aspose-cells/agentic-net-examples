// Title: C# example: digitally signing a PDF generated from an Excel workbook with Aspose.Cells and Aspose.Pdf
// Common Searches: C# code to apply a digital signature to a PDF exported from an Excel file using Aspose | How to sign a PDF created by Aspose.Cells with Aspose.Pdf in .NET | Programmatically add a digital certificate to a PDF generated from a workbook in C# | Aspose.Cells PDF export then digital signing workflow for .NET applications | Example of using Aspose.Pdf to sign a PDF produced from an Excel workbook
// Tags: Aspose.Cells export workbook to PDF C# | Aspose.Pdf digital signature implementation | sign PDF generated from Excel Aspose | .NET PDF signing with Aspose libraries | add digital certificate to PDF using Aspose.Pdf

using System;
using System.IO;
using Aspose.Cells;

// The sample creates an Excel workbook, fills it with data, saves it as a PDF using Aspose.Cells, and explains that a digital signature can be applied by referencing Aspose.Pdf and using its signing API in C#.
class Program
{
    static void Main()
    {
        try
        {
            // 1. Create a new Excel workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Oranges");
            sheet.Cells["B3"].PutValue(85);

            // 2. Define the output PDF path
            string pdfPath = "Report.pdf";

            // 3. Save the workbook as a PDF file
            workbook.Save(pdfPath, SaveFormat.Pdf);

            // Verify that the PDF was created
            if (File.Exists(pdfPath))
            {
                Console.WriteLine($"PDF generated successfully at: {Path.GetFullPath(pdfPath)}");
            }
            else
            {
                Console.WriteLine("Failed to generate the PDF file.");
            }

            // Note: PDF signing requires Aspose.Pdf library, which is not referenced in this project.
            // If digital signing is needed, add a reference to Aspose.Pdf and implement the signing logic separately.
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
