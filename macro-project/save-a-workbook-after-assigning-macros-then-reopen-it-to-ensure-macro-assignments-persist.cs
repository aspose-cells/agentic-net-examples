using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsMacroDemo
{
    public class MacroPersistenceDemo
    {
        public static void Run()
        {
            try
            {
                // Path for the temporary macro-enabled workbook
                string tempMacroPath = "TempMacroWorkbook.xlsm";

                // 1. Create a new workbook and save it as a macro-enabled file (XLSM)
                Workbook wb = new Workbook();
                wb.Save(tempMacroPath, SaveFormat.Xlsm); // Creates the VBA project container

                // Ensure the temporary file exists before loading
                if (!File.Exists(tempMacroPath))
                    throw new FileNotFoundException($"File not found: {tempMacroPath}");

                // 2. Load the saved workbook to access its VBA project
                Workbook macroWb = new Workbook(tempMacroPath);

                // 3. Add a new VBA module and set its code
                int moduleIndex = macroWb.VbaProject.Modules.Add(VbaModuleType.Class, "DemoModule");
                VbaModule module = macroWb.VbaProject.Modules[moduleIndex];
                module.Codes = "Sub HelloWorld()\r\n    MsgBox \"Hello from VBA!\"\r\nEnd Sub";

                // 4. Save the workbook again with the macro attached
                string macroWorkbookPath = "WorkbookWithMacro.xlsm";
                macroWb.Save(macroWorkbookPath, SaveFormat.Xlsm);

                // Ensure the macro workbook exists before verification
                if (!File.Exists(macroWorkbookPath))
                    throw new FileNotFoundException($"File not found: {macroWorkbookPath}");

                // 5. Reopen the workbook to verify that the macro persists
                Workbook verifyWb = new Workbook(macroWorkbookPath);
                Console.WriteLine("HasMacro after reload: " + verifyWb.HasMacro); // Should be True

                // Optional: Clean up temporary file
                if (File.Exists(tempMacroPath))
                {
                    File.Delete(tempMacroPath);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            MacroPersistenceDemo.Run();
        }
    }
}