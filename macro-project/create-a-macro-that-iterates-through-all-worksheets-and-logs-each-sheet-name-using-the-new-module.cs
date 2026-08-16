// Title: C# – Add a VBA module that logs worksheet names and save as .xlsm with Aspose.Cells
// Description: A C# example that creates an empty workbook, injects a VBA module named LogSheets containing a Sub that iterates through all worksheets and outputs each sheet name via Debug.Print, then saves the file as a macro‑enabled .xlsm using Aspose.Cells.
// Keywords: Aspose.Cells VBA module | C# add VBA to workbook | log worksheet names VBA | save workbook as xlsm | iterate worksheets macro | .xlsm generation Aspose
// Common Searches: how to add a VBA module with Aspose.Cells .NET | C# code to inject VBA that lists sheet names | save macro‑enabled workbook using Aspose.Cells | Aspose.Cells example for logging worksheet names | create .xlsm file programmatically C#
// Developer Intent: Inject a VBA module that loops through every worksheet, logs each name, and produce a macro‑enabled workbook.
// Use Cases: Generate template workbooks that automatically list their sheets for debugging or audit trails. | Distribute Excel files with a built‑in macro that end users can run to view all worksheet names without writing code. | Provide a starter VBA module that can be customized and embedded in Excel reports created by Aspose.Cells.
// AI Prompts: Generate C# code using Aspose.Cells to add a VBA module named LogSheets that prints each worksheet name and saves the workbook as .xlsm. | Show how to modify the injected VBA to write sheet names to a new worksheet instead of using Debug.Print. | Explain how to add multiple VBA modules to an Aspose.Cells workbook and call one from another in C#.

using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsMacroExample
{
    // A C# example that creates an empty workbook, injects a VBA module named LogSheets containing a Sub that iterates through all worksheets and outputs each sheet name via Debug.Print, then saves the file as a macro‑enabled .xlsm using Aspose.Cells.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (empty workbook with a default worksheet)
            Workbook workbook = new Workbook();

            // Add a standard VBA module to the workbook
            // Using VbaModuleType.Class as an example; adjust the type if needed
            int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Class, "LogSheets");
            VbaModule module = workbook.VbaProject.Modules[moduleIndex];

            // VBA code that iterates through all worksheets and logs each sheet name
            string vbaCode = @"
Sub LogSheetNames()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Debug.Print ws.Name
    Next ws
End Sub
";
            module.Codes = vbaCode;

            // Save the workbook in a macro‑enabled format
            workbook.Save("WorkbookWithLogSheetsMacro.xlsm", SaveFormat.Xlsm);
        }
    }
}
