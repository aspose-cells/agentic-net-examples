// Title: C# – List Worksheet Paper Size Mode, FitToPagesWide & FitToPagesTall with Aspose.Cells
// Description: Loads a workbook, iterates through every worksheet, reads the PageSetup to determine whether the paper size is automatic or a specific enum, and prints the paper‑size mode together with FitToPagesWide and FitToPagesTall values. The workbook is then saved, demonstrating the full Aspose.Cells lifecycle for .NET.
// Keywords: Aspose.Cells | C# worksheet print settings | PaperSize mode Aspose.Cells | FitToPagesWide | FitToPagesTall | PageSetup API | Aspose.Cells .NET | print scaling report
// Common Searches: Aspose.Cells get worksheet paper size | How to read FitToPagesWide in Aspose.Cells | Print settings report for all sheets Aspose.Cells | Retrieve page setup properties C# Aspose.Cells
// Developer Intent: Generate a console report that shows each worksheet’s paper‑size mode and its FitToPagesWide / FitToPagesTall values using Aspose.Cells for .NET.
// Use Cases: Verify print scaling before exporting to PDF | Audit workbook print configuration for batch printing | Document worksheet layout settings for stakeholders | Ensure compliance with corporate printing standards
// AI Prompts: Provide C# code that writes the worksheet name, paper size mode, FitToPagesWide and FitToPagesTall to a CSV file using Aspose.Cells. | Extend the sample to include orientation, margins, and header/footer details in the output. | Create a reusable method that returns a Dictionary<string, (string PaperMode, int FitWide, int FitTall)> for a given workbook.

using System;
using Aspose.Cells;

// Loads a workbook, iterates through every worksheet, reads the PageSetup to determine whether the paper size is automatic or a specific enum, and prints the paper‑size mode together with FitToPagesWide and FitToPagesTall values. The workbook is then saved, demonstrating the full Aspose.Cells lifecycle for .NET.
class WorksheetPrintSettingsReport
{
    static void Main()
    {
        // Load an existing workbook (adjust the path as needed)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Iterate through all worksheets and output their print settings
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            Worksheet ws = workbook.Worksheets[i];
            PageSetup ps = ws.PageSetup;

            // Determine whether the paper size is automatic or a specific enum value
            string paperMode = ps.IsAutomaticPaperSize ? "Automatic" : ps.PaperSize.ToString();

            Console.WriteLine($"Worksheet: {ws.Name}");
            Console.WriteLine($"  Paper Size Mode : {paperMode}");
            Console.WriteLine($"  FitToPagesWide  : {ps.FitToPagesWide}");
            Console.WriteLine($"  FitToPagesTall  : {ps.FitToPagesTall}");
            Console.WriteLine();
        }

        // Save the workbook (demonstrates the required lifecycle handling)
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
