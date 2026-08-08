// Title: Save an Aspose.Cells workbook with slicers and VBA macros to a macro‑enabled XLSM file (C#)
// Description: Load an existing .xlsm workbook that contains slicers and VBA macros, optionally verify the HasMacro flag, and save it as a macro‑enabled XLSM file while preserving all slicer objects and VBA code using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# save XLSM | preserve slicers | Workbook.HasMacro | macro‑enabled Excel | VBA macros | SaveFormat.Xlsm | Excel slicer retention
// Common Searches: Aspose.Cells save workbook with slicers as .xlsm | C# preserve VBA macros when saving Excel with Aspose.Cells | How to keep slicer objects after saving to macro‑enabled format | Check for macros before saving Excel file using Aspose.Cells | SaveFormat.Xlsm example with slicers
// Developer Intent: Save a workbook that contains slicers and VBA macros as a macro‑enabled XLSM file using Aspose.Cells for .NET.
// Use Cases: Programmatically modify a template workbook that includes slicers and macros, then export it as a new .xlsm file. | Validate the presence of VBA code before deciding to save in macro‑enabled format. | Batch‑process multiple Excel files, ensuring slicer functionality and macros remain intact after each save.
// AI Prompts: Generate C# code with Aspose.Cells to load an .xlsm file, check Workbook.HasMacro, and save it preserving slicers and VBA macros. | Show how to use SaveFormat.Xlsm to retain slicer objects when exporting a workbook with macros. | Explain why Aspose.Cells keeps slicer definitions intact when saving to a macro‑enabled Excel file.

using System;
using Aspose.Cells;

namespace AsposeCellsMacroEnabledSave
{
    // Load an existing .xlsm workbook that contains slicers and VBA macros, optionally verify the HasMacro flag, and save it as a macro‑enabled XLSM file while preserving all slicer objects and VBA code using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook that contains slicers and VBA macros.
            // Replace "input.xlsm" with the path to your source file.
            Workbook workbook = new Workbook("input.xlsm");

            // Verify that the workbook indeed contains macros (optional).
            if (workbook.HasMacro)
            {
                Console.WriteLine("Workbook contains VBA macros. Saving as macro‑enabled file...");
            }
            else
            {
                Console.WriteLine("Warning: Workbook does not contain macros.");
            }

            // Save the workbook to a macro‑enabled Excel format (XLSM) preserving the VBA code.
            // The Save method with (string, SaveFormat) follows the provided lifecycle rule.
            workbook.Save("output.xlsm", SaveFormat.Xlsm);

            Console.WriteLine("Workbook saved successfully as macro‑enabled file: output.xlsm");
        }
    }
}
