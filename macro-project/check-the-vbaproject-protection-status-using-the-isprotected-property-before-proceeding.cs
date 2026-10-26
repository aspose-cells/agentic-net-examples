// Title: Determine if an Excel workbook's VBA project is protected using Aspose.Cells IsProtected property in C#
// Common Searches: C# Aspose.Cells check VBA project IsProtected before modifying macros | how to detect protected VBA project in an .xlsx file using Aspose.Cells | load password‑protected workbook with Aspose.Cells and verify VBA project protection | prevent processing of protected VBA macros in C# with Aspose.Cells | Aspose.Cells IsProtected property example for VBA project in C#
// Tags: Aspose.Cells VBA protection check | C# IsProtected VBA project | load encrypted workbook Aspose.Cells | skip processing protected VBA macros | detect protected VBA project C#

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file with Aspose.Cells, prompts for a password if the workbook is encrypted, checks for a VBA project, reads its IsProtected flag, and aborts further processing when the VBA project is protected.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Attempt to load the workbook without a password
            workbook = new Workbook(inputPath);
        }
        catch (CellsException ex) when (ex.Message.Contains("Please provide password"))
        {
            // Prompt the user for a password if the workbook is protected
            Console.Write("The workbook is password protected. Please enter the password: ");
            string password = Console.ReadLine();

            try
            {
                var loadOptions = new LoadOptions { Password = password };
                workbook = new Workbook(inputPath, loadOptions);
            }
            catch (Exception innerEx)
            {
                Console.WriteLine($"Failed to load workbook with the provided password: {innerEx.Message}");
                return;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Verify that the workbook contains a VBA project
        if (workbook?.VbaProject != null)
        {
            // Check the protection status of the VBA project
            bool isProtected = workbook.VbaProject.IsProtected;
            Console.WriteLine($"VBA project protected: {isProtected}");

            // If the VBA project is protected, stop further processing
            if (isProtected)
            {
                Console.WriteLine("The VBA project is protected. Exiting.");
                return;
            }

            // Place additional processing code here for unprotected VBA projects
            // ...
        }
        else
        {
            Console.WriteLine("No VBA project found in the workbook.");
        }
    }
}
