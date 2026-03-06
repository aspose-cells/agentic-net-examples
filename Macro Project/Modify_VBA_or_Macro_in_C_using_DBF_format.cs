using System;
using Aspose.Cells;
using Aspose.Cells.Loading;
using Aspose.Cells.Saving;
using Aspose.Cells.Vba;

public class ModifyVbaInDbfDemo
{
    public static void Run()
    {
        // Load the DBF file with default load options
        DbfLoadOptions loadOptions = new DbfLoadOptions();
        Workbook dbfWorkbook = new Workbook("input.dbf", loadOptions);

        // Create a new workbook that will hold the data and the VBA macro
        Workbook macroWorkbook = new Workbook();

        // Copy all data from the loaded DBF worksheet to the new workbook
        Worksheet srcSheet = dbfWorkbook.Worksheets[0];
        Worksheet destSheet = macroWorkbook.Worksheets[0];
        int rowCount = srcSheet.Cells.MaxDataRow + 1; // include header row
        destSheet.Cells.CopyRows(srcSheet.Cells, 0, 0, rowCount);

        // Add a VBA module to the workbook and set its code
        int moduleIndex = macroWorkbook.VbaProject.Modules.Add(VbaModuleType.Class, "MyMacro");
        VbaModule vbaModule = macroWorkbook.VbaProject.Modules[moduleIndex];
        vbaModule.Codes =
            "Sub ShowData()\n" +
            "    MsgBox \"Data loaded from DBF file\"\n" +
            "End Sub";

        // Save the workbook as a macro‑enabled Excel file
        macroWorkbook.Save("output.xlsm", SaveFormat.Xlsm);

        // Optionally, save the data back to DBF with ExportAsString enabled
        DbfSaveOptions dbfSaveOptions = new DbfSaveOptions
        {
            ExportAsString = true
        };
        macroWorkbook.Save("output.dbf", dbfSaveOptions);
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        ModifyVbaInDbfDemo.Run();
    }
}