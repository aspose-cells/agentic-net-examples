// Title: C# – Protect Each Excel Worksheet with a Distinct Password Using Aspose.Cells
// Description: The sample loads a workbook, derives a base key from the source file name, builds a separate password for every worksheet by appending its index, applies full sheet protection (ProtectionType.All) with that key, and writes the secured file to the target path.
// Keywords: Aspose.Cells C# worksheet protection | Excel sheet password per sheet | distinct sheet passwords .NET | protect workbook sheets individually | file‑name based password generation | ProtectionType.All usage | C# Excel security example | batch workbook protection Aspose | Excel encryption .NET | Aspose.Cells API password
// Common Searches: C# protect individual Excel sheets with Aspose.Cells | generate different passwords for each worksheet programmatically | use file name to create worksheet passwords in .NET | apply ProtectionType.All to all sheets Aspose | save workbook after sheet protection C# | how to set per‑sheet password in Aspose.Cells | Excel sheet security using C# and Aspose
// Developer Intent: Add per‑sheet password protection to an Excel workbook via Aspose.Cells.
// Use Cases: Distribute a multi‑department workbook where each department's tab is locked with its own credential. | Comply with data‑privacy policies by encrypting every sheet before sending the file to external partners. | Automate processing of dozens of workbooks, assigning sheet‑specific passwords derived from each file's name.
// AI Prompts: Write C# code with Aspose.Cells that secures every worksheet using a password that combines the workbook name and sheet index. | Show how to replace the file‑name‑based password pattern with a custom format, such as "Dept_{SheetName}_{Date}". | Explain how to programmatically verify the password of a particular worksheet after it has been protected with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The sample loads a workbook, derives a base key from the source file name, builds a separate password for every worksheet by appending its index, applies full sheet protection (ProtectionType.All) with that key, and writes the secured file to the target path.
    class ProtectWorksheets
    {
        public static void Run(string inputFilePath, string outputFilePath)
        {
            // Verify input file exists
            if (!File.Exists(inputFilePath))
                throw new FileNotFoundException($"Input file not found: {inputFilePath}");

            // Load the existing workbook
            Workbook workbook = new Workbook(inputFilePath);

            // Base password derived from the file name (without extension)
            string basePassword = Path.GetFileNameWithoutExtension(inputFilePath);

            // Protect each worksheet with a unique password
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Create a unique password for the current worksheet
                string sheetPassword = $"{basePassword}_Sheet{sheet.Index}";

                // Protect the worksheet with all protection types using the unique password
                sheet.Protect(ProtectionType.All, sheetPassword, null);
            }

            // Save the workbook with protected worksheets
            workbook.Save(outputFilePath);
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Expecting two arguments: input file path and output file path
                if (args.Length < 2)
                {
                    Console.WriteLine("Usage: AsposeCellsExamples <inputFilePath> <outputFilePath>");
                    return;
                }

                string inputPath = args[0];
                string outputPath = args[1];

                ProtectWorksheets.Run(inputPath, outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine($"File error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
