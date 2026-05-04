using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaMhtExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Ensure the workbook has a VBA project by saving as a macro-enabled file and reloading
            string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsm");
            workbook.Save(tempPath, SaveFormat.Xlsm);
            workbook = new Workbook(tempPath);
            File.Delete(tempPath);

            // Add a new procedural VBA module named "MhtModule"
            int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Procedural, "MhtModule");
            VbaModule vbaModule = workbook.VbaProject.Modules[moduleIndex];

            // Load the MHT source file content (create a dummy file if it does not exist)
            string mhtFilePath = "module.mht";
            if (!File.Exists(mhtFilePath))
            {
                string dummyMht = "Content-Type: multipart/related; boundary=\"----=_NextPart_000_0000\"\r\n\r\n------=_NextPart_000_0000\r\nContent-Type: text/html; charset=\"utf-8\"\r\n\r\n<html><body><p>Dummy MHT content</p></body></html>\r\n------=_NextPart_000_0000--";
                File.WriteAllText(mhtFilePath, dummyMht);
            }
            string mhtContent = File.ReadAllText(mhtFilePath);

            // Assign the MHT content to the module's code
            vbaModule.Codes = mhtContent;

            // Save the workbook with the VBA module (lifecycle: save)
            string outputPath = "WorkbookWithMhtModule.xlsm";
            workbook.Save(outputPath, SaveFormat.Xlsm);

            Console.WriteLine($"Workbook saved to '{outputPath}' with VBA module containing MHT source.");
        }
    }
}