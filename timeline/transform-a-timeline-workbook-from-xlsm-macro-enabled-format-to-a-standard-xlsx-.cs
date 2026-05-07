using System;
using System.IO;
using Aspose.Cells;

class TimelineWorkbookConverter
{
    static void Main()
    {
        // Source macro‑enabled workbook (XLSM)
        string sourcePath = Path.Combine(Environment.CurrentDirectory, "timeline.xlsm");

        // Destination workbook without macros (XLSX)
        string destPath = Path.Combine(Environment.CurrentDirectory, "timeline.xlsx");

        // Ensure the source file exists; create a placeholder if it does not
        if (!File.Exists(sourcePath))
        {
            Workbook placeholder = new Workbook();
            placeholder.Save(sourcePath, SaveFormat.Xlsm);
        }

        // Load the XLSM file
        Workbook workbook = new Workbook(sourcePath);

        // Remove all VBA/macros from the workbook
        workbook.RemoveMacro();

        // Save the cleaned workbook as a standard XLSX file
        workbook.Save(destPath, SaveFormat.Xlsx);

        Console.WriteLine($"Workbook converted successfully to: {destPath}");
    }
}