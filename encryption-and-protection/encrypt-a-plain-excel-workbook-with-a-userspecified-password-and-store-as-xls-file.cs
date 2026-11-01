// Title: Encrypt an Excel workbook with a password and save it as an XLS file using Aspose.Cells for .NET
// Common Searches: C# how to encrypt an existing Excel file and output as .xls using Aspose.Cells | Aspose.Cells set password on workbook before saving to Excel 97‑2003 format | Encrypt Excel workbook with user supplied password via command line arguments in .NET | Save password‑protected workbook as XLS with Aspose.Cells API | Apply workbook password protection and convert to XLS in C#
// Tags: Workbook.Settings.Password encryption | Save as Excel97To2003 format | C# password protect XLS workbook | Aspose.Cells command line encryption | Encrypt Excel file for Excel 97‑2003 output

using System;
using Aspose.Cells;

// The example reads an input path, output path, and password from the command line, loads the source workbook with Aspose.Cells, assigns the password via Workbook.Settings.Password, and saves the encrypted workbook as an Excel 97‑2003 (.xls) file using SaveFormat.Excel97To2003.
class Program
{
    static void Main(string[] args)
    {
        // Expect three arguments: input file path, output file path, and password
        if (args.Length < 3)
        {
            Console.WriteLine("Usage: Program.exe <inputPath> <outputPath> <password>");
            return;
        }

        string inputPath = args[0];
        string outputPath = args[1];
        string password = args[2];

        // Load the plain Excel workbook (any supported format)
        Workbook workbook = new Workbook(inputPath);

        // Apply password protection (encryption) to the workbook
        workbook.Settings.Password = password;

        // Save the encrypted workbook as an XLS file (Excel 97‑2003 format)
        workbook.Save(outputPath, SaveFormat.Excel97To2003);
    }
}
