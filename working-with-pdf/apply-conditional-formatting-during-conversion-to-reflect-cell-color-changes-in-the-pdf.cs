// Title: Apply Conditional Formatting and Keep Cell Colors When Converting Excel to PDF with Aspose.Cells for .NET (C#)
// Description: This C# example creates a workbook, fills column A with numeric values, adds three conditional‑formatting rules (greater than 50, equal to 50, less than 50) that set light‑green or light‑coral backgrounds, configures PdfSaveOptions (MergeAreas = true, CalculateFormula = true) and saves the sheet as a PDF so the conditional‑format colors are rendered in the output.
// Keywords: Aspose.Cells | C# | conditional formatting | PDF conversion | preserve cell colors | PdfSaveOptions | MergeAreas | CalculateFormula | Excel to PDF | color highlighting | range A1:A10 | light green background | light coral background
// Common Searches: Aspose.Cells keep conditional formatting colors in PDF | C# export Excel to PDF with background colors | PdfSaveOptions MergeAreas effect | how to preserve conditional formatting when saving as PDF using Aspose.Cells | conditional formatting thresholds Excel to PDF C# | Aspose.Cells PDF conversion color loss fix
// Developer Intent: The developer wants to apply conditional‑formatting rules to cells and generate a PDF that displays the same background colors defined by those rules.
// Use Cases: Generate a PDF report that highlights values above a threshold in green and below in coral. | Create compliance‑ready financial statements where key figures are automatically colored during Excel‑to‑PDF conversion. | Automate batch processing of Excel worksheets to PDFs while retaining conditional‑formatting for visual analytics. | Produce printable dashboards with color‑coded performance metrics directly from code.
// AI Prompts: Show how to add gradient conditional formatting based on numeric ranges and keep the gradient visible in the exported PDF. | Demonstrate date‑based conditional formatting (e.g., overdue dates in red) and export the workbook to PDF with Aspose.Cells. | Explain how to achieve the same color preservation without using MergeAreas, using alternative rendering options. | Provide code to apply icon sets conditional formatting and retain icons in the PDF output. | Suggest ways to optimize PDF size while preserving conditional‑formatting colors for large workbooks.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// This C# example creates a workbook, fills column A with numeric values, adds three conditional‑formatting rules (greater than 50, equal to 50, less than 50) that set light‑green or light‑coral backgrounds, configures PdfSaveOptions (MergeAreas = true, CalculateFormula = true) and saves the sheet as a PDF so the conditional‑format colors are rendered in the output.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Fill sample numeric data in column A (A1:A10)
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 0].PutValue(i * 10); // 0,10,20,...,90
            }

            // Add conditional formatting to the range A1:A10
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

            // Define the area for the conditional formatting (A1:A10)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 9,
                StartColumn = 0,
                EndColumn = 0
            };
            fcc.AddArea(area);

            // Condition 1: values > 50 -> light green background
            int condIdxGt = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "50", null);
            FormatCondition fcGt = fcc[condIdxGt];
            fcGt.Style.BackgroundColor = Color.LightGreen;

            // Condition 2: values = 50 -> light green background (to emulate >= 50)
            int condIdxEq = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.Equal, "50", null);
            FormatCondition fcEq = fcc[condIdxEq];
            fcEq.Style.BackgroundColor = Color.LightGreen;

            // Condition 3: values < 50 -> light coral background
            int condIdxLt = fcc.AddCondition(FormatConditionType.CellValue, OperatorType.LessThan, "50", null);
            FormatCondition fcLt = fcc[condIdxLt];
            fcLt.Style.BackgroundColor = Color.LightCoral;

            // Configure PDF save options to preserve conditional formatting colors
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                MergeAreas = true,          // Merge conditional formatting before rendering
                CalculateFormula = true    // Ensure any formulas are evaluated
            };

            // Define output file path
            string outputPath = "ConditionalFormattingDemo.pdf";

            // Save the workbook as PDF; the cell background colors set by conditional formatting will appear in the PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
