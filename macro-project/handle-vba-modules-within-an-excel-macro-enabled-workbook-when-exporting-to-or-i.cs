using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.Ods;

class VbaOdsExample
{
    static void Main()
    {
        // Paths for the source macro-enabled workbook and output files
        string macroWorkbookPath = "TemplateWithMacro.xlsm";
        string odsPath = "ExportedFile.ods";
        string macroOnlyPath = "MacroOnlyCopy.xlsm";

        // Ensure the source macro-enabled workbook exists; if not, create a simple one with a macro
        if (!File.Exists(macroWorkbookPath))
        {
            Workbook wb = new Workbook();
            // Add a simple VBA module
            int modIdx = wb.VbaProject.Modules.Add(VbaModuleType.Procedural, "InitModule");
            VbaModule mod = wb.VbaProject.Modules[modIdx];
            mod.Codes = "Sub InitMacro()\n    MsgBox \"Initial macro\"\nEnd Sub";
            wb.Save(macroWorkbookPath, SaveFormat.Xlsm);
        }

        // Load the macro-enabled workbook
        Workbook macroWb = new Workbook(macroWorkbookPath);

        // Verify that the workbook contains VBA macros
        if (macroWb.HasMacro)
        {
            Console.WriteLine("Source workbook contains VBA macros.");

            // ------------------------------------------------------------
            // Preserve the VBA project by copying it to a new workbook
            // ------------------------------------------------------------
            Workbook copyWb = new Workbook();

            // Create a temporary macro-enabled file to initialise the VBA project
            copyWb.Save("temp.xlsm", SaveFormat.Xlsm);
            copyWb = new Workbook("temp.xlsm");
            File.Delete("temp.xlsm");

            // Copy the VBA project from the source workbook
            copyWb.VbaProject.Copy(macroWb.VbaProject);
            copyWb.Save(macroOnlyPath, SaveFormat.Xlsm);
            Console.WriteLine($"Copied VBA project saved to {macroOnlyPath}");

            // ------------------------------------------------------------
            // Remove macros before exporting to ODS (ODS does not support VBA)
            // ------------------------------------------------------------
            macroWb.RemoveMacro();

            // Configure ODS save options
            OdsSaveOptions odsOptions = new OdsSaveOptions
            {
                GeneratorType = OdsGeneratorType.LibreOffice,
                IsStrictSchema11 = true,
                OdfStrictVersion = OpenDocumentFormatVersionType.Odf12
            };

            // Save the workbook as ODS
            macroWb.Save(odsPath, odsOptions);
            Console.WriteLine($"Workbook exported to ODS at {odsPath}");
        }
        else
        {
            Console.WriteLine("Source workbook does not contain VBA macros.");
        }

        // ------------------------------------------------------------
        // Load the ODS file (macros are not present) and add a new VBA module
        // ------------------------------------------------------------
        Workbook odsWb = new Workbook(odsPath);

        // Initialise a VBA project by saving as a temporary .xlsm and reloading
        string tempXlsm = "TempFromOds.xlsm";
        odsWb.Save(tempXlsm, SaveFormat.Xlsm);
        Workbook macroEnabledFromOds = new Workbook(tempXlsm);
        File.Delete(tempXlsm);

        // Add a new procedural VBA module
        int newModuleIdx = macroEnabledFromOds.VbaProject.Modules.Add(VbaModuleType.Procedural, "NewModule");
        VbaModule newModule = macroEnabledFromOds.VbaProject.Modules[newModuleIdx];
        newModule.Codes = "Sub HelloFromOds()\n    MsgBox \"Hello from ODS imported workbook!\"\nEnd Sub";

        // Save the workbook with the newly added macro
        string finalMacroPath = "FinalWithNewMacro.xlsm";
        macroEnabledFromOds.Save(finalMacroPath, SaveFormat.Xlsm);
        Console.WriteLine($"Final workbook with new VBA module saved to {finalMacroPath}");
    }
}