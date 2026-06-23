using System;
using Aspose.Cells;

class DetectDefaultEncryption
{
    static void Main(string[] args)
    {
        // Verify that a file path was supplied
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: DetectDefaultEncryption <workbookPath>");
            return;
        }

        string workbookPath = args[0];

        // Load the workbook (no password assumed)
        Workbook workbook = new Workbook(workbookPath);

        // Check if the workbook is encrypted
        if (workbook.Settings.IsEncrypted)
        {
            // Determine whether the default encryption algorithm is used
            if (workbook.Settings.IsDefaultEncrypted)
            {
                Console.WriteLine("Workbook uses the default encryption algorithm. Consider upgrading to a stronger encryption method.");
            }
            else
            {
                Console.WriteLine("Workbook is encrypted with a custom algorithm.");
            }
        }
        else
        {
            Console.WriteLine("Workbook is not encrypted.");
        }

        // Clean up resources
        workbook.Dispose();
    }
}