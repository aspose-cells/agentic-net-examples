using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Path to the text file that contains the VBA project password
        string passwordFilePath = "vba_password.txt";

        // If the password file does not exist, create it with a default password
        if (!File.Exists(passwordFilePath))
        {
            File.WriteAllText(passwordFilePath, "MySecretPassword");
        }

        // Read the password from the text file (trim to remove any newline characters)
        string vbaPassword = File.ReadAllText(passwordFilePath).Trim();

        // Create a new workbook (this will also create a VBA project container)
        Workbook workbook = new Workbook();

        // Protect the VBA project.
        // The first argument (true) locks the project for viewing.
        // The second argument is the password read from the txt file.
        workbook.VbaProject.Protect(true, vbaPassword);

        // Save the workbook as a macro‑enabled file so the VBA project is retained
        workbook.Save("ProtectedVbaProject.xlsm", SaveFormat.Xlsm);
    }
}