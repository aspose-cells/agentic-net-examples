// Title: Saving a modified Excel workbook as an XLSM file while preserving VBA macros with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx workbook, injects VBA modules, and saves it as a macro‑enabled .xlsm using Aspose.Cells. | Show how to use Aspose.Cells SaveFormat.Xlsm to keep an existing VBA project when exporting a workbook.
// Common Searches: Aspose.Cells C# how to export workbook to macro‑enabled XLSM format | retain VBA code after modifying Excel file with Aspose.Cells | C# example saving workbook with macros using Aspose.Cells SaveFormat.Xlsm | convert .xlsx to .xlsm while keeping macros in .NET | Aspose.Cells preserve VBA project when saving workbook
// Tags: Aspose.Cells SaveFormat.Xlsm implementation | create macro‑enabled XLSM file with C# | add VBA code to workbook via Aspose.Cells | maintain VBA macros on workbook export | generate XLSM from modified workbook Aspose.Cells

using Aspose.Cells;

// The C# program loads an existing Excel file (input.xlsx) with Aspose.Cells, applies modifications and adds VBA code, then saves the workbook as output.xlsm using SaveFormat.Xlsm to retain the VBA project.
class Program
{
    static void Main()
    {
        // Load an existing workbook (or create a new one)
        Workbook workbook = new Workbook("input.xlsx");

        // ... perform modifications and add VBA code to the workbook ...

        // Save the workbook as XLSM to preserve the VBA project
        workbook.Save("output.xlsm", SaveFormat.Xlsm);
    }
}
