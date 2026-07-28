// Title: C# – Unprotect a Worksheet, Edit a Named Range, and Re‑protect with Password using Aspose.Cells
// Description: Loads an existing workbook, removes worksheet protection with a password, accesses the named range "MyRange", updates its first cell, reapplies protection with the same password, and saves the file. Demonstrates safe editing of protected sheets in Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | unprotect worksheet | protect worksheet password | named range edit | worksheet protection | modify named range Aspose.Cells | Aspose.Cells workbook protection example
// Common Searches: Aspose.Cells unprotect worksheet C# | How to edit a named range in a protected sheet using Aspose.Cells | Re‑apply worksheet protection after modifying cells Aspose.Cells .NET | C# code to unprotect, change, and protect Excel sheet with Aspose.Cells | Aspose.Cells example for named range modification in protected workbook
// Developer Intent: Programmatically remove worksheet protection, change data in a named range, and restore protection with the original password.
// Use Cases: Automated update of configuration values stored in a protected named range before distribution. | Data‑correction routine that temporarily lifts protection, edits a specific range, and re‑secures the sheet. | Generating reports that require temporary access to protected cells, then re‑applying security.
// AI Prompts: Write C# code with Aspose.Cells to unprotect a worksheet using a password, modify the cells of a named range called "MyRange", and protect the worksheet again with the same password. | Explain how to handle a missing named range when editing a protected workbook with Aspose.Cells in C#. | Show how to protect a worksheet with specific ProtectionType options while preserving the existing password using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an existing workbook, removes worksheet protection with a password, accesses the named range "MyRange", updates its first cell, reapplies protection with the same password, and saves the file. Demonstrates safe editing of protected sheets in Aspose.Cells for .NET.
    public class UnprotectModifyProtectDemo
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";
                const string password = "pwd";

                // Verify input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Unprotect the worksheet using the password
                sheet.Unprotect(password);

                // Retrieve the named range "MyRange"
                Name namedRange = workbook.Worksheets.Names["MyRange"];
                if (namedRange == null)
                {
                    Console.WriteLine("Named range 'MyRange' does not exist.");
                    return;
                }

                // Get the actual cell range that the name refers to
                Aspose.Cells.Range range = namedRange.GetRange();

                // Example modification: set the first cell of the range to a new value
                range[0, 0].PutValue("Modified Value");

                // Re‑apply protection with the same password (oldPassword not required, pass null)
                sheet.Protect(ProtectionType.All, password, null);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
