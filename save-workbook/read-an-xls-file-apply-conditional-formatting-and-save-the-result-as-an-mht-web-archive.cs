// Title: C# – Load XLS, Add >100 Conditional Formatting, and Export to MHT with Aspose.Cells
// Description: This example loads an XLS workbook (or creates a new one), defines a yellow‑fill rule for cells in A1:D10 whose value exceeds 100, ensures the target folder exists, and saves the result as an MHTML web‑archive using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# example | conditional formatting XLS | highlight values greater than 100 | save workbook as MHT | MHTML export .NET | web archive conversion | C# Excel to MHTML | Aspose.Cells SaveFormat.MHtml | programmatic style rule | Excel to web archive
// Common Searches: Aspose.Cells add conditional formatting and export to MHT | C# convert XLS to MHTML with formatting | how to highlight cells >100 using Aspose.Cells | save Excel workbook as .mht file in .NET | create web‑archive from Excel with Aspose
// Developer Intent: Apply a value‑based style to a range and generate an MHTML version of the workbook.
// Use Cases: Produce browser‑friendly reports that flag high numbers before distribution. | Automate email attachments by converting styled spreadsheets to a single‑file web archive. | Batch‑process legacy XLS files, apply a uniform highlight rule, and archive each as MHT for compliance.
// AI Prompts: Generate C# code that loads an XLS, adds a yellow background to cells >100 in A1:D10, and saves the file as MHTML using Aspose.Cells. | Explain how to modify the formatted range or change the fill color before exporting to a web archive. | Show best practices for error handling and creating the output directory when converting a workbook to MHT.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsConditionalFormattingToMht
{
    // This example loads an XLS workbook (or creates a new one), defines a yellow‑fill rule for cells in A1:D10 whose value exceeds 100, ensures the target folder exists, and saves the result as an MHTML web‑archive using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xls";
                const string outputPath = "output.mht";

                // Load existing workbook if the file exists; otherwise create a new workbook.
                Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

                // Access the first worksheet.
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a new ConditionalFormatting object to the worksheet.
                int cfIndex = worksheet.ConditionalFormattings.Add();
                var cf = worksheet.ConditionalFormattings[cfIndex];

                // Define the range for conditional formatting (A1:D10) and associate it.
                CellArea area = CellArea.CreateCellArea("A1", "D10");
                cf.AddArea(area);

                // Add a conditional formatting rule: highlight cells with value > 100.
                // The AddCondition method requires a second formula argument; pass null when not needed.
                int conditionIndex = cf.AddCondition(
                    FormatConditionType.CellValue,
                    OperatorType.GreaterThan,
                    "100",
                    null);

                // Retrieve the created rule.
                FormatCondition condition = cf[conditionIndex];

                // Define the style to apply when the condition is met (yellow background).
                Style style = workbook.CreateStyle();
                style.ForegroundColor = Color.Yellow;
                style.Pattern = BackgroundType.Solid;

                // Assign the style to the condition.
                condition.Style = style;

                // Ensure the output directory exists.
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as an MHTML (web archive) file.
                workbook.Save(outputPath, SaveFormat.MHtml);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
