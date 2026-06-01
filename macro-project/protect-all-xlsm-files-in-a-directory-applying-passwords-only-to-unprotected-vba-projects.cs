using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class ProtectVbaProjects
{
    static void Main(string[] args)
    {
        // Directory containing the XLSM files. Change as needed or pass as first argument.
        string folderPath = args.Length > 0 ? args[0] : @"C:\ExcelFiles";

        // Password to apply to unprotected VBA projects.
        const string vbaPassword = "MyVbaPassword";

        // Get all .xlsm files in the specified directory.
        string[] files = Directory.GetFiles(folderPath, "*.xlsm", SearchOption.TopDirectoryOnly);

        foreach (string filePath in files)
        {
            // Load the workbook.
            Workbook workbook = new Workbook(filePath);

            // Check if the VBA project is already protected.
            if (!workbook.VbaProject.IsProtected)
            {
                // Protect the VBA project without locking it for viewing.
                workbook.VbaProject.Protect(false, vbaPassword);

                // Save the workbook, overwriting the original file.
                workbook.Save(filePath, SaveFormat.Xlsm);
                Console.WriteLine($"Protected VBA project in: {Path.GetFileName(filePath)}");
            }
            else
            {
                Console.WriteLine($"VBA project already protected in: {Path.GetFileName(filePath)}");
            }

            // Release resources.
            workbook.Dispose();
        }
    }
}