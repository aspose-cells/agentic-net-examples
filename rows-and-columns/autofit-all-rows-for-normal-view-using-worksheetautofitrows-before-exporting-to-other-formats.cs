// Title: Auto-fit all rows in every worksheet with Worksheet.AutoFitRows before exporting to PDF using Aspose.Cells for .NET
// Common Searches: c# aspnet autosize row heights in all worksheets before saving as pdf with aspose.cells | how to apply Worksheet.AutoFitRows to each sheet in a workbook using Aspose.Cells .NET | auto fit rows for normal view in Excel file before converting to PDF with Aspose.Cells | Aspose.Cells .NET iterate worksheets and auto fit rows prior to export
// Tags: Worksheet.AutoFitRows batch processing | auto fit rows before PDF conversion | adjust row heights Aspose.Cells .NET | apply auto fit to all worksheets | export workbook to PDF after row auto fit

using Aspose.Cells;

// The example creates or loads an Excel workbook, loops through every worksheet calling Worksheet.AutoFitRows to size rows based on their content, and then saves the workbook to a PDF file.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        var workbook = new Workbook(); // create rule

        // If you need to load an existing file, use the load rule instead:
        // workbook = new Workbook("input.xlsx"); // load rule

        // Auto‑fit all rows in every worksheet for normal view
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.AutoFitRows(); // fits rows based on their content
        }

        // Export to the desired format (e.g., PDF, XLSX, etc.)
        workbook.Save("output.pdf", SaveFormat.Pdf); // save rule
    }
}
