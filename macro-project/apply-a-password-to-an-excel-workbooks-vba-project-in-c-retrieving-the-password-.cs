using System;
using System.IO;
using Aspose.Cells;

class ApplyVbaPassword
{
    static void Main()
    {
        // Paths to the Excel file, TSV file containing the password, and the output file
        string workbookPath = "input.xlsm";
        string tsvPath = "passwords.tsv";
        string outputPath = "output.xlsm";

        // Verify that the required files exist
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Workbook file not found: {workbookPath}");
            return;
        }

        if (!File.Exists(tsvPath))
        {
            Console.WriteLine($"TSV file not found: {tsvPath}");
            return;
        }

        // Retrieve the password from the first column of the first line in the TSV file
        string password = null;
        using (var reader = new StreamReader(tsvPath))
        {
            string line = reader.ReadLine();
            if (!string.IsNullOrEmpty(line))
            {
                var parts = line.Split('\t');
                if (parts.Length > 0)
                    password = parts[0];
            }
        }

        if (string.IsNullOrEmpty(password))
        {
            Console.WriteLine("Password not found in TSV file.");
            return;
        }

        // Load the workbook (macro-enabled) that contains a VBA project
        Workbook workbook = new Workbook(workbookPath);

        // Protect the VBA project and lock it for viewing using the retrieved password
        if (workbook.VbaProject != null)
        {
            workbook.VbaProject.Protect(true, password);
        }
        else
        {
            Console.WriteLine("The workbook does not contain a VBA project.");
            return;
        }

        // Save the workbook as a macro-enabled file
        workbook.Save(outputPath, SaveFormat.Xlsm);

        Console.WriteLine($"VBA project protected with password and saved to {outputPath}");
    }
}