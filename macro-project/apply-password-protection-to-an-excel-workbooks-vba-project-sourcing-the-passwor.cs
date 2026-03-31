using System;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Vba;

class ProtectVbaProjectFromCsv
{
    static void Main()
    {
        // Path to the CSV file containing the password (first column of the first row)
        string csvPath = "password.csv";

        // Read the password from the CSV file
        string password = "";
        if (File.Exists(csvPath))
        {
            // Assuming the password is the first value in the first line
            string firstLine = File.ReadLines(csvPath).FirstOrDefault();
            if (!string.IsNullOrEmpty(firstLine))
            {
                password = firstLine.Split(',')[0].Trim();
            }
        }

        // Create a new workbook (lifecycle create rule)
        Workbook workbook = new Workbook();

        // Save temporarily as a macro‑enabled workbook to ensure a VBA project exists
        string tempPath = Path.Combine(Path.GetTempPath(), "temp.xlsm");
        workbook.Save(tempPath, SaveFormat.Xlsm);

        // Load the temporary workbook (lifecycle load rule)
        workbook = new Workbook(tempPath);

        // Protect the VBA project using the password read from CSV
        // islockedForViewing set to false (protect without locking for viewing)
        workbook.VbaProject.Protect(false, password);

        // Save the final workbook with protected VBA project (lifecycle save rule)
        string outputPath = "protected_vba.xlsm";
        workbook.Save(outputPath, SaveFormat.Xlsm);

        // Clean up the temporary file
        if (File.Exists(tempPath))
        {
            File.Delete(tempPath);
        }

        Console.WriteLine($"VBA project protected and saved to '{outputPath}'.");
    }
}