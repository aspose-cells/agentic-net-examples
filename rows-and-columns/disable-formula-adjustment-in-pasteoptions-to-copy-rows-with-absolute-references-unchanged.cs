// Title: Copy worksheet rows while preserving absolute formula references using PasteOptions in Aspose.Cells for .NET
// AI Prompts: Write C# code that copies rows from a source worksheet to a destination worksheet with PasteOptions.PasteType set to Formulas so that $A$1 style absolute references stay unchanged. | Show how to call Cells.CopyRows with a CopyOptions instance and a PasteOptions configured for formula‑only pasting to avoid adjusting relative references. | Generate a complete example that copies two rows containing mixed absolute and relative formulas, disables value and format copying, and saves the workbook as an .xlsx file.
// Common Searches: Aspose.Cells copy rows without changing absolute cell references in C# | prevent formula adjustment when pasting formulas using PasteOptions in Aspose.Cells | copy only formulas between worksheets preserving $A$1 references Aspose.Cells .NET | C# example of CopyRows with PasteType.Formulas and no value formatting
// Tags: CopyRows method with formula‑only PasteOptions Aspose.Cells | preserve $A$1 references during worksheet row transfer .NET | disable formula adjustment using PasteOptions Aspose.Cells | formula‑only paste without values Aspose.Cells | C# Aspose.Cells example for row transfer

using System;
using Aspose.Cells;

// // This program creates a workbook, adds a source sheet with absolute and relative formulas, then copies the first two rows to a destination sheet using Cells.CopyRows with PasteOptions.PasteType = Formulas, ensuring absolute references remain unchanged, and saves the result as FormulaCopy_NoAdjustment.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the source worksheet
        Workbook workbook = new Workbook();
        Worksheet srcSheet = workbook.Worksheets[0];
        srcSheet.Name = "Source";

        // Populate source sheet with data and formulas
        // Row 1 – absolute references (will stay unchanged after copy)
        srcSheet.Cells["A1"].PutValue(10);
        srcSheet.Cells["B1"].PutValue(20);
        srcSheet.Cells["C1"].Formula = "=$A$1+$B$1";

        // Row 2 – relative references (will be adjusted normally)
        srcSheet.Cells["A2"].PutValue(30);
        srcSheet.Cells["B2"].PutValue(40);
        srcSheet.Cells["C2"].Formula = "=A1+B1";

        // Create a destination worksheet
        Worksheet dstSheet = workbook.Worksheets.Add("Destination");

        // Prepare copy and paste options
        // CopyOptions – default behavior
        CopyOptions copyOptions = new CopyOptions();

        // PasteOptions – copy only formulas; this disables any value/format copying
        // Absolute references in the formulas remain exactly as they are
        PasteOptions pasteOptions = new PasteOptions
        {
            PasteType = PasteType.Formulas,
            // No additional adjustments are required; absolute references stay unchanged
        };

        // Copy the first two rows from source to destination using the overload that accepts both options
        // Parameters: source cells, source start row, destination start row, number of rows, copy options, paste options
        dstSheet.Cells.CopyRows(srcSheet.Cells, 0, 0, 2, copyOptions, pasteOptions);

        // Save the workbook
        workbook.Save("FormulaCopy_NoAdjustment.xlsx");
    }
}
