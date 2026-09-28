// Title: How to convert an XLSX workbook to PDF using Aspose.Cells in C# (.NET)
// AI Prompts: Write C# code that loads an Excel .xlsx file with Aspose.Cells and saves it as a PDF using SaveFormat.Pdf. | Show how to batch convert multiple .xlsx files to PDFs in a C# loop using Aspose.Cells. | Explain how to configure PDF export options (e.g., page layout, image quality) when saving a workbook with Aspose.Cells in .NET.
// Common Searches: Aspose.Cells C# convert single Excel file to PDF without opening Excel | programmatically export .xlsx to PDF using Aspose.Cells SaveFormat.Pdf | C# code sample for saving workbook as PDF with Aspose.Cells .NET library | how to set PDF export options when converting Excel to PDF with Aspose.Cells | convert Excel workbook to PDF in .NET Core using Aspose.Cells
// Tags: Aspose.Cells xlsx to pdf conversion C# | Workbook.Save SaveFormat.Pdf usage | Aspose.Cells PDF export settings C# | Load workbook Aspose.Cells C# example | PDF generation from Excel Aspose.Cells .NET

using Aspose.Cells;

// The example loads an XLSX workbook from a specified path using Aspose.Cells and saves it as a PDF file by calling Workbook.Save with the SaveFormat.Pdf option.
class Program
{
    static void Main()
    {
        // Load the XLSX workbook from the specified file path
        Workbook workbook = new Workbook("input.xlsx");

        // Save the loaded workbook as a PDF document
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
