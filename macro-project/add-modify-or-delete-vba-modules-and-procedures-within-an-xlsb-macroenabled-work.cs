using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaDemo
{
    public class VbaOperations
    {
        public static void Run()
        {
            // -----------------------------------------------------------------
            // 1. Create a new workbook (XLSB supports macros)
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();

            // -----------------------------------------------------------------
            // 2. Ensure the workbook has a VBA project.
            //    For a newly created workbook, the VbaProject is present but empty.
            // -----------------------------------------------------------------
            VbaProject vbaProject = workbook.VbaProject;

            // -----------------------------------------------------------------
            // 3. Add a new procedural module named "UtilityModule"
            // -----------------------------------------------------------------
            int utilityModuleIndex = vbaProject.Modules.Add(VbaModuleType.Procedural, "UtilityModule");
            VbaModule utilityModule = vbaProject.Modules[utilityModuleIndex];

            // Set VBA code for the new module
            utilityModule.Codes =
                "Public Sub ShowMessage()\r\n" +
                "    MsgBox \"Hello from UtilityModule!\"\r\n" +
                "End Sub";

            // -----------------------------------------------------------------
            // 4. Add a worksheet‑specific module (code behind the first sheet)
            // -----------------------------------------------------------------
            Worksheet sheet = workbook.Worksheets[0];
            int sheetModuleIndex = vbaProject.Modules.Add(sheet);
            VbaModule sheetModule = vbaProject.Modules[sheetModuleIndex];

            // Set VBA code for the worksheet module
            sheetModule.Codes =
                "Private Sub Worksheet_SelectionChange(ByVal Target As Range)\r\n" +
                "    MsgBox \"Cell \" & Target.Address & \" selected.\"\r\n" +
                "End Sub";

            // -----------------------------------------------------------------
            // 5. Modify an existing module's code (if it exists)
            //    Here we demonstrate updating the code of \"UtilityModule\"
            // -----------------------------------------------------------------
            VbaModule existingModule = vbaProject.Modules["UtilityModule"];
            existingModule.Codes =
                "Public Sub ShowMessage()\r\n" +
                "    MsgBox \"Updated message from UtilityModule!\"\r\n" +
                "End Sub";

            // -----------------------------------------------------------------
            // 6. Delete a module by name (remove the worksheet module)
            // -----------------------------------------------------------------
            // The worksheet module name equals the sheet's CodeName
            string sheetModuleName = sheet.CodeName;
            vbaProject.Modules.Remove(sheetModuleName);

            // -----------------------------------------------------------------
            // 7. Optionally remove all macros from the workbook
            // -----------------------------------------------------------------
            // Uncomment the following line to strip all VBA code from the file
            // workbook.RemoveMacro();

            // -----------------------------------------------------------------
            // 8. Save the workbook as a macro‑enabled XLSB file
            // -----------------------------------------------------------------
            string outputPath = "VbaDemoWorkbook.xlsb";
            workbook.Save(outputPath, SaveFormat.Xlsb);

            Console.WriteLine($"Workbook saved to '{outputPath}'.");
            Console.WriteLine($"HasMacro: {workbook.HasMacro}");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            VbaOperations.Run();
        }
    }
}