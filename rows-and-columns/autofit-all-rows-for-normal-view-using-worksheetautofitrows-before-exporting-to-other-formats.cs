// Title: Automatically adjust row heights with Worksheet.AutoFitRows and export the workbook to PDF using Aspose.Cells in C#
// AI Prompts: Generate C# code that loads an Aspose.Cells workbook, invokes Worksheet.AutoFitRows to adjust all row heights for normal view, and then saves the workbook as a PDF file. | Show how to apply Worksheet.AutoFitRows on a selected worksheet and subsequently export the same workbook to PDF and XLSX formats within a .NET application.
// Common Searches: c# aspose.cells auto fit rows before saving as pdf | how to use Worksheet.AutoFitRows to adjust row height in Aspose.Cells | export aspose.cells workbook to pdf after auto fitting rows | auto fit all rows in excel file using Aspose.Cells .NET | aspose.cells auto fit rows normal view pdf export example
// Tags: Aspose.Cells Worksheet.AutoFitRows implementation | auto-fit rows prior to PDF export Aspose.Cells | row height normalization C# Aspose.Cells | PDF generation after row auto-fit Aspose.Cells | C# workbook export with adjusted row heights

// Create a new workbook
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();

// Access the first worksheet (or any specific worksheet)
Aspose.Cells.Worksheet worksheet = workbook.Worksheets[0];

// Auto‑fit all rows in the worksheet for normal view
worksheet.AutoFitRows();

// Export the workbook to the desired format (e.g., PDF, XLSX, etc.)
// Example: save as PDF
workbook.Save("ExportedDocument.pdf", Aspose.Cells.SaveFormat.Pdf);
