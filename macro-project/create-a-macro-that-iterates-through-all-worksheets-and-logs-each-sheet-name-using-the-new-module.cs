using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a new VBA module to the workbook (procedural type)
        int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Procedural, "LogSheets");
        VbaModule module = workbook.VbaProject.Modules[moduleIndex];

        // VBA macro that iterates through all worksheets and logs each sheet name
        string vbaCode =
            "Sub LogAllSheetNames()\n" +
            "    Dim ws As Worksheet\n" +
            "    For Each ws In ThisWorkbook.Worksheets\n" +
            "        Debug.Print ws.Name\n" +
            "    Next ws\n" +
            "End Sub";

        // Assign the VBA code to the module
        module.Codes = vbaCode;

        // Save the workbook as a macro‑enabled file
        workbook.Save("LogSheetsMacro.xlsm", SaveFormat.Xlsm);
    }
}