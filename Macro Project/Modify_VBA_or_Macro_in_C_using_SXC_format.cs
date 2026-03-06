using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsMacroExample
{
    class Program
    {
        static void Main(string[] args)
        {
            ModifyVbaAndSaveAsSxc.Run();
        }
    }

    public class ModifyVbaAndSaveAsSxc
    {
        public static void Run()
        {
            string inputPath = "input.xlsm";

            Workbook workbook = new Workbook(inputPath);

            if (workbook.HasMacro)
            {
                workbook.RemoveMacro();
            }

            VbaProject vbaProject = workbook.VbaProject;

            int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Class, "NewModule");

            VbaModule vbaModule = vbaProject.Modules[moduleIndex];
            vbaModule.Codes = "Sub NewMacro()\r\n    MsgBox \"Hello from new macro!\"\r\nEnd Sub";

            string outputPath = "output.sxc";
            workbook.Save(outputPath, SaveFormat.Sxc);

            Console.WriteLine($"Workbook processed and saved as SXC: {outputPath}");
        }
    }
}