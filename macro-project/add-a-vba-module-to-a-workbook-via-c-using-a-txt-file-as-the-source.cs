using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaModuleFromTxt
{
    class Program
    {
        static void Main()
        {
            // Path to the text file containing VBA code
            string txtPath = "MacroCode.txt";

            // Ensure the source file exists
            if (!File.Exists(txtPath))
            {
                Console.WriteLine($"Source file not found: {txtPath}");
                return;
            }

            // Read VBA code from the text file
            string vbaCode = File.ReadAllText(txtPath);

            // Create a new workbook (macro‑enabled will be saved later)
            Workbook workbook = new Workbook();

            // Add a new procedural VBA module named "MyMacroModule"
            int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Procedural, "MyMacroModule");

            // Retrieve the added module and assign the code read from the file
            VbaModule module = workbook.VbaProject.Modules[moduleIndex];
            module.Codes = vbaCode;

            // Save the workbook as a macro‑enabled file
            string outputPath = "WorkbookWithVba.xlsm";
            workbook.Save(outputPath, SaveFormat.Xlsm);

            Console.WriteLine($"Workbook saved with VBA module at: {outputPath}");
        }
    }
}