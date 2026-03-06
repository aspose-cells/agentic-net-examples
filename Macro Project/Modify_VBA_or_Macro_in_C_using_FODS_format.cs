using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaFodsDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.fods";

            // Ensure the input file exists; if not, create a simple workbook and save it as FODS
            if (!File.Exists(inputPath))
            {
                var tempWb = new Workbook();
                tempWb.Worksheets[0].Name = "Sheet1";
                tempWb.Worksheets[0].Cells["A1"].PutValue("Sample Data");
                tempWb.Save(inputPath, SaveFormat.Fods);
            }

            // Load the workbook from the FODS file
            var workbook = new Workbook(inputPath);

            // If the workbook already contains a macro, remove it; otherwise, add a new VBA module
            if (workbook.HasMacro)
            {
                Console.WriteLine("Workbook already contains a macro. Removing existing macro...");
                workbook.RemoveMacro();
            }
            else
            {
                Console.WriteLine("Workbook does not contain a macro. Adding a new VBA module...");

                // Save temporarily as macro‑enabled file to initialize the VBA project
                string tempMacroPath = "temp.xlsm";
                workbook.Save(tempMacroPath, SaveFormat.Xlsm);

                // Reload the workbook with the VBA project initialized
                workbook = new Workbook(tempMacroPath);

                // Add a new class module to the VBA project
                int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Class, "MyMacroModule");

                // Retrieve the newly added module and set its VBA code
                VbaModule module = workbook.VbaProject.Modules[moduleIndex];
                module.Codes = @"Sub HelloWorld()
    MsgBox ""Hello from VBA!""
End Sub";

                // Delete the temporary macro‑enabled file
                File.Delete(tempMacroPath);
            }

            // Save the modified workbook back to FODS format
            string outputPath = "output.fods";
            workbook.Save(outputPath, SaveFormat.Fods);

            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
    }
}