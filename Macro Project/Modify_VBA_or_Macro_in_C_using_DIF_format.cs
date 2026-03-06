using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.Loading;

namespace AsposeCellsMacroModification
{
    public class ModifyVbaUsingDif
    {
        public static void Run()
        {
            // Path to the source DIF file (if it exists)
            string difPath = "TemplateData.dif";

            Workbook workbook;

            if (File.Exists(difPath))
            {
                // Load the DIF file with default load options
                DifLoadOptions loadOptions = new DifLoadOptions();
                workbook = new Workbook(difPath, loadOptions);
            }
            else
            {
                // Create a new workbook if the DIF file does not exist
                workbook = new Workbook();
            }

            // Ensure the workbook has a VBA project (required for macro manipulation)
            if (!workbook.HasMacro)
            {
                // Save temporarily as macro-enabled workbook to initialize the VBA project
                string tempXlsm = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsm");
                workbook.Save(tempXlsm, SaveFormat.Xlsm);

                // Reload the temporary file which now contains an empty VBA project
                workbook = new Workbook(tempXlsm);
                File.Delete(tempXlsm);
            }

            // Add a new class module (or retrieve an existing one) to the VBA project
            int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Class, "MyMacroModule");
            VbaModule vbaModule = workbook.VbaProject.Modules[moduleIndex];

            // Set or replace the VBA code in the module
            vbaModule.Codes =
                "Sub HelloWorld()\n" +
                "    MsgBox \"Hello from modified VBA!\"\n" +
                "End Sub";

            // OPTIONAL: Save the workbook back to DIF format (macros are not stored in DIF, this is just for demonstration)
            DifSaveOptions difSaveOptions = new DifSaveOptions
            {
                ClearData = true,
                CreateDirectory = true,
                RefreshChartCache = true
            };
            workbook.Save("ModifiedData.dif", difSaveOptions);

            // Save the workbook as a macro‑enabled Excel file so the VBA code is retained
            string outputXlsm = "ModifiedWorkbook.xlsm";
            workbook.Save(outputXlsm, SaveFormat.Xlsm);

            Console.WriteLine($"Workbook saved with modified VBA macro: {outputXlsm}");
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            ModifyVbaUsingDif.Run();
        }
    }
}