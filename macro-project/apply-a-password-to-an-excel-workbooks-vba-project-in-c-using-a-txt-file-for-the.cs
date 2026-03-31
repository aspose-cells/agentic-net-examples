using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class ApplyVbaPassword
{
    static void Main()
    {
        // Path to the text file that contains the VBA password
        string passwordFilePath = "vba_password.txt";

        // Read the password from the file (remove any trailing newline characters)
        string vbaPassword = File.ReadAllText(passwordFilePath).Trim();

        // Path to the source workbook (must be a macro‑enabled file)
        string sourceWorkbookPath = "input.xlsm";

        Workbook workbook;

        // Load the workbook if it exists; otherwise create a new macro‑enabled workbook
        if (File.Exists(sourceWorkbookPath))
        {
            workbook = new Workbook(sourceWorkbookPath);
        }
        else
        {
            // Create a new workbook and save it as .xlsm to initialise a VBA project
            workbook = new Workbook();
            workbook.Save("temp.xlsm", SaveFormat.Xlsm);
            workbook = new Workbook("temp.xlsm");
            File.Delete("temp.xlsm");
        }

        // Protect the VBA project and lock it for viewing using the password read from the file
        workbook.VbaProject.Protect(true, vbaPassword);

        // Save the workbook with the protected VBA project
        string outputWorkbookPath = "output_protected.xlsm";
        workbook.Save(outputWorkbookPath, SaveFormat.Xlsm);
    }
}