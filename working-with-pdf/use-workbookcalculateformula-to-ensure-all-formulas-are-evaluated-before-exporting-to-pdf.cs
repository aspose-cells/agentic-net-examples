// Title: Calculate workbook formulas with Aspose.Cells for .NET before exporting to PDF
// AI Prompts: Generate C# code that loads an .xlsx file, runs Workbook.CalculateFormula, and saves the result as a PDF using Aspose.Cells. | Show how to force formula recalculation in a workbook prior to PDF conversion with Aspose.Cells in a .NET application. | Provide a .NET snippet that ensures all Excel formulas are evaluated before calling workbook.Save with SaveFormat.Pdf.
// Common Searches: Aspose.Cells C# calculate formulas before PDF conversion | how to force formula recalculation when saving Excel as PDF using Aspose.Cells | Workbook.CalculateFormula usage example for PDF export in .NET | evaluate all formulas in an Excel workbook prior to PDF generation with Aspose.Cells | C# code to recalculate Excel formulas then convert to PDF using Aspose.Cells
// Tags: calculate formulas before PDF export | evaluate Excel formulas Aspose.Cells C# | force formula recalculation for PDF conversion | Aspose.Cells PDF generation with updated cells | recalculate workbook prior to SaveFormat.Pdf

using Aspose.Cells;

// The example loads an Excel workbook, forces evaluation of every formula with Workbook.CalculateFormula, and then saves the workbook as a PDF file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the workbook from an existing Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Ensure all formulas are evaluated
        workbook.CalculateFormula();

        // Export the workbook to PDF format
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
