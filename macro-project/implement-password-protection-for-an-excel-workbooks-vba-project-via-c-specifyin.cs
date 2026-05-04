using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaProtectionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default is .xlsx)
            Workbook wb = new Workbook();

            // Add a worksheet to ensure the workbook has at least one sheet
            wb.Worksheets.Add();

            // Save as a macro‑enabled workbook to create a VBA project container
            string tempPath = Path.Combine(Path.GetTempPath(), "temp.xlsm");
            wb.Save(tempPath, SaveFormat.Xlsm);

            // Reload the workbook so that the VBA project is initialized
            Workbook macroWb = new Workbook(tempPath);

            // Define the password in HTML format
            string htmlPassword = "<html><body>MySecurePassword</body></html>";

            // Protect the VBA project and lock it for viewing
            // islockedForViewing = true, password = htmlPassword
            macroWb.VbaProject.Protect(true, htmlPassword);

            // Save the protected workbook
            string outputPath = "VbaProjectProtected.xlsm";
            macroWb.Save(outputPath, SaveFormat.Xlsm);

            // Clean up temporary file
            if (File.Exists(tempPath))
                File.Delete(tempPath);

            Console.WriteLine($"Workbook saved with VBA project protected. Path: {outputPath}");
        }
    }
}