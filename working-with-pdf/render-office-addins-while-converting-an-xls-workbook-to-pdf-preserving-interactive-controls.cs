// Title: C# – Convert XLS with Office Add‑Ins to PDF while preserving interactive controls using Aspose.Cells
// Description: Loads an XLS workbook containing Office Add‑Ins (form controls, check boxes, drop‑downs) and uses Aspose.Cells.Utility.ConversionUtility to create a PDF that retains those interactive elements.
// Keywords: Aspose.Cells | C# PDF conversion | XLS to PDF | preserve form controls | Office Add‑Ins | interactive PDF | ConversionUtility | retain interactivity | Excel to PDF with controls
// Common Searches: Aspose.Cells keep Excel form controls in PDF | C# convert XLS with add‑ins to PDF | preserve interactive elements when converting Excel to PDF | how to retain check boxes in PDF using Aspose.Cells | convert Excel workbook with Office Add‑Ins to PDF .NET
// Developer Intent: Generate a PDF from an XLS workbook that contains Office Add‑Ins, ensuring the embedded form controls stay functional.
// Use Cases: Create printable PDFs of templates that still allow users to check boxes or select options. | Archive Excel reports with active controls so reviewers can interact with the data in a read‑only format. | Distribute Excel‑based questionnaires as PDFs while preserving dropdowns and radio buttons.
// AI Prompts: Provide C# code that uses Aspose.Cells ConversionUtility to convert an XLS with form controls to a PDF that keeps the controls interactive. | Explain how Aspose.Cells handles Office Add‑Ins during Excel‑to‑PDF conversion and note any limitations. | Give a step‑by‑step guide to verify that interactive controls are retained after converting an Excel workbook to PDF with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Loads an XLS workbook containing Office Add‑Ins (form controls, check boxes, drop‑downs) and uses Aspose.Cells.Utility.ConversionUtility to create a PDF that retains those interactive elements.
class Program
{
    static void Main()
    {
        // Path to the source Excel workbook (XLS) that contains Office Add‑Ins (e.g., form controls)
        string sourcePath = "input.xls";

        // Path where the resulting PDF will be saved
        string destPath = "output.pdf";

        // Convert the Excel file to PDF.
        // ConversionUtility handles loading the workbook, preserving interactive controls,
        // and saving the output in PDF format.
        ConversionUtility.Convert(sourcePath, destPath);

        Console.WriteLine("Excel workbook has been successfully converted to PDF.");
    }
}
