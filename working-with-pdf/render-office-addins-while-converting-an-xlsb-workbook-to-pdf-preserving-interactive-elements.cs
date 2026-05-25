using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Source XLSB workbook (contains Office Add‑Ins / macros)
        string sourcePath = "input.xlsb";

        // Destination PDF file
        string destPath = "output.pdf";

        // Load the XLSB workbook (create/load rule)
        Workbook workbook = new Workbook(sourcePath);

        // Save the workbook as PDF (save rule). 
        // Aspose.Cells preserves interactive elements such as macros/add‑ins during PDF conversion.
        workbook.Save(destPath, SaveFormat.Pdf);
    }
}