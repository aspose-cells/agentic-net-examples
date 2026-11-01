// Title: Recalculate all formulas in an Excel workbook with Aspose.Cells for .NET before exporting to PDF
// AI Prompts: Load an .xlsx file, invoke Workbook.CalculateFormula to refresh every formula, and then save the workbook as a PDF using Aspose.Cells in C#. | Force a full formula evaluation on a workbook and generate a PDF output with the Aspose.Cells .NET API. | Update all calculated cells in an Excel workbook before converting it to PDF with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# recalculate formulas before PDF export | How to use Workbook.CalculateFormula and then save as PDF in .NET | Force formula evaluation in Excel file with Aspose.Cells and convert to PDF | C# example for updating Excel formulas and exporting to PDF using Aspose.Cells
// Tags: Workbook.CalculateFormula C# | Excel to PDF conversion with refreshed formulas | Aspose.Cells formula recalculation before PDF export | C# Aspose.Cells update calculated cells | PDF export after Excel formula evaluation

// Load an existing workbook (replace with your actual file path)
var workbook = new Aspose.Cells.Workbook("input.xlsx");

// Recalculate all formulas in the workbook
workbook.CalculateFormula();

// Export the workbook to PDF (replace with your desired output path)
workbook.Save("output.pdf", Aspose.Cells.SaveFormat.Pdf);
