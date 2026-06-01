using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample worksheets (optional, for demonstration)
        workbook.Worksheets[0].Name = "Sheet1";
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");

        // Add a new VBA module to the workbook
        VbaProject vbaProject = workbook.VbaProject;
        int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Procedural, "LogModule");
        VbaModule module = vbaProject.Modules[moduleIndex];

        // VBA macro that iterates through all worksheets and logs each sheet name
        string vbaCode = @"
Sub LogSheetNames()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Debug.Print ws.Name
    Next ws
End Sub
";
        module.Codes = vbaCode;

        // Save the workbook as a macro‑enabled file
        workbook.Save("LogSheetNames.xlsm", SaveFormat.Xlsm);
    }
}