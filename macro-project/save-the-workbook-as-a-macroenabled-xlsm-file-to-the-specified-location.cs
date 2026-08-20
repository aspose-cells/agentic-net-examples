// Title: Save a Workbook as Macro‑Enabled XLSM with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a new Workbook, write sample data, and persist it as a macro‑enabled XLSM file using Aspose.Cells' Save method with SaveFormat.Xlsm, including a console confirmation of the saved path.
// Keywords: Aspose.Cells | C# | SaveFormat.Xlsm | macro enabled workbook | XLSM file | save workbook as XLSM | Aspose.Cells example | Excel macro file | C# Aspose.Cells save
// Common Searches: Aspose.Cells save workbook as XLSM C# | How to create macro enabled Excel file with Aspose.Cells | C# example for saving XLSM using Aspose.Cells | SaveFormat.Xlsm Aspose.Cells tutorial | Generate XLSM file programmatically in .NET
// Developer Intent: Create and persist a macro‑enabled Excel workbook (XLSM) using Aspose.Cells in C#.
// Use Cases: Produce a report that includes VBA macros for downstream users. | Automate generation of a template workbook with macro support for further customization. | Export application data to an XLSM file that can be opened directly in Excel with functional macros.
// AI Prompts: Generate C# code with Aspose.Cells that creates a workbook, adds data, and saves it as a macro‑enabled XLSM file. | Explain how to attach a VBA project to an Aspose.Cells workbook before saving it as XLSM. | Show how to load an existing .vba project into a workbook and persist it as a macro‑enabled file using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsMacroSaveDemo
{
    // Demonstrates how to create a new Workbook, write sample data, and persist it as a macro‑enabled XLSM file using Aspose.Cells' Save method with SaveFormat.Xlsm, including a console confirmation of the saved path.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and add sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Macro Enabled Workbook");
            sheet.Cells["B1"].PutValue(DateTime.Now);

            // Define the output file path (macro‑enabled XLSM)
            string outputPath = "MacroEnabledWorkbook.xlsm";

            // Save the workbook as XLSM using the Save method with SaveFormat.Xlsm
            workbook.Save(outputPath, SaveFormat.Xlsm);

            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
    }
}
