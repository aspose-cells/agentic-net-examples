using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Define the print area covering columns A‑D and rows 10‑30
        worksheet.PageSetup.PrintArea = "A10:D30";

        // Save the workbook as a PDF file
        workbook.Save("CustomPrintArea.pdf", SaveFormat.Pdf);
    }
}