using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class MacroWorkbookCheck
{
    static void Main()
    {
        // -------------------------------------------------
        // 1. Create a workbook and ensure it has a VBA project
        // -------------------------------------------------
        Workbook wb = new Workbook();

        // Save as a macro‑enabled file to a temporary stream and reload
        // to initialize the VbaProject (VbaProject is read‑only until a macro file is loaded)
        using (MemoryStream tempStream = new MemoryStream())
        {
            wb.Save(tempStream, SaveFormat.Xlsm);
            tempStream.Position = 0;
            wb = new Workbook(tempStream);
        }

        // -------------------------------------------------
        // 2. Add a VBA module to the workbook
        // -------------------------------------------------
        int moduleIndex = wb.VbaProject.Modules.Add(VbaModuleType.Class, "TestModule");
        VbaModule module = wb.VbaProject.Modules[moduleIndex];
        module.Codes = "Sub Test()\r\n    MsgBox \"Hello from VBA!\"\r\nEnd Sub";

        // -------------------------------------------------
        // 3. Save the macro‑enabled workbook to a memory stream
        // -------------------------------------------------
        MemoryStream macroStream = new MemoryStream();
        wb.Save(macroStream, SaveFormat.Xlsm);
        macroStream.Position = 0; // Reset for reading

        // -------------------------------------------------
        // 4. Load the workbook from the memory stream
        // -------------------------------------------------
        Workbook loadedWb = new Workbook(macroStream);

        // -------------------------------------------------
        // 5. Verify that the loaded workbook contains at least one VBA module
        // -------------------------------------------------
        bool hasAtLeastOneModule = loadedWb.VbaProject != null && loadedWb.VbaProject.Modules.Count > 0;
        Console.WriteLine("Workbook contains VBA module: " + hasAtLeastOneModule);
    }
}