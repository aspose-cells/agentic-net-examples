// Title: How to set worksheet page zoom to 80 % with Aspose.Cells for .NET and export the sheet as a PDF
// AI Prompts: Set the PageSetup.Zoom property of a worksheet to 80 % and save the workbook as a PDF using Aspose.Cells for .NET. | Apply an 80 % page scaling to the first worksheet and generate a PDF file with Aspose.Cells in C#.
// Common Searches: Aspose.Cells .NET set worksheet page zoom before PDF conversion | C# change Excel sheet zoom to 80 percent and export to PDF with Aspose.Cells | how to adjust PageSetup.Zoom when saving a workbook as PDF using Aspose.Cells
// Tags: PageSetup.Zoom 80 percent | worksheet page scaling Aspose.Cells | export worksheet to PDF .NET | custom zoom PDF conversion Aspose.Cells | adjust page zoom before PDF save

using Aspose.Cells;

// // Creates a new workbook, sets the first worksheet's page zoom to 80 %, and saves the workbook as a PDF file named output.pdf.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set the page zoom to 80%
        sheet.PageSetup.Zoom = 80;

        // Export the worksheet to PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
