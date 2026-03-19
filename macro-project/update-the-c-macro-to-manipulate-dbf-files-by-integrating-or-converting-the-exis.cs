using System;
using Aspose.Cells;
using Aspose.Cells.Saving;
using Aspose.Cells.Vba;

public class DbfMacroIntegrationDemo
{
    public static void Main(string[] args)
    {
        Run();
    }

    public static void Run()
    {
        // Load a macro‑enabled workbook (create/load lifecycle)
        Workbook workbook = new Workbook("TemplateWithMacro.xlsm");

        // Check if the workbook contains VBA macros
        if (workbook.HasMacro)
        {
            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Iterate through all VBA modules and modify their code
            for (int i = 0; i < vbaProject.Modules.Count; i++)
            {
                VbaModule module = vbaProject.Modules[i];
                // Append a comment indicating that the macro logic has been migrated to .NET
                module.Codes += "\n' Converted to .NET handling for DBF export";
            }
        }

        // ------------------------------------------------------------
        // Example: replicate a simple VBA calculation in .NET code
        // ------------------------------------------------------------
        Worksheet sheet = workbook.Worksheets[0];

        // Assume the original VBA summed values in column B (index 1)
        int lastDataRow = sheet.Cells.MaxDataRow;
        double sum = 0;

        for (int row = 1; row <= lastDataRow; row++) // start from row 2 (skip header)
        {
            object cellValue = sheet.Cells[row, 1].Value; // column B
            if (cellValue is double d)
                sum += d;
        }

        // Write the result back to the worksheet (mimicking VBA output)
        sheet.Cells[lastDataRow + 1, 0].PutValue("Total");
        sheet.Cells[lastDataRow + 1, 1].PutValue(sum);

        // DBF format does not support macros; remove them before saving
        workbook.RemoveMacro();

        // Configure DBF save options (export all values as strings)
        DbfSaveOptions saveOptions = new DbfSaveOptions
        {
            ExportAsString = true
        };

        // Save the workbook as a DBF file (save lifecycle)
        workbook.Save("Result.dbf", saveOptions);
    }
}