// Title: Automatically adjust worksheet row heights to fit cell content with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads a workbook, calls the AutoFitRows method on a worksheet, and saves the file. | Show how to enable automatic row‑height scaling for all rows in an Aspose.Cells worksheet using .NET. | Create a minimal program that creates a new Excel file, applies AutoFitRows to the first sheet, and outputs output.xlsx.
// Common Searches: Aspose.Cells C# how to auto fit row height for entire worksheet | Set Excel rows to automatically adjust height based on content using Aspose.Cells .NET | AutoFitRows method example for adjusting row heights in a .xlsx file | C# code to auto‑scale row heights in Aspose.Cells workbook | Make worksheet rows auto‑height in Aspose.Cells for .NET
// Tags: row height auto‑fit with Aspose.Cells | C# auto‑fit worksheet rows | Excel row height auto‑adjust Aspose.Cells | auto‑fit rows based on cell content .NET | set default row height to auto Aspose.Cells

using Aspose.Cells;

// // This program creates a new workbook, retrieves the first worksheet, automatically adjusts all row heights to fit their content using AutoFitRows, and saves the result as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Automatically adjust the height of all rows based on their cell content
        sheet.AutoFitRows();

        // Save the workbook
        workbook.Save("output.xlsx");
    }
}
