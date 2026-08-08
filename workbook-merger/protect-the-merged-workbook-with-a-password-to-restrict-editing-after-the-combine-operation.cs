// Title: C# – Password‑protect a merged Excel workbook with Aspose.Cells
// Description: Loads a workbook created by a combine operation, applies full workbook protection (structure and windows) using Workbook.Protect with a password, and saves the result as a password‑protected Excel file.
// Keywords: Aspose.Cells | .NET | C# | Workbook.Protect | password protection | merged workbook | Excel file | protect structure | protect windows | save protected workbook | combine workbooks
// Common Searches: How to add password protection to a combined Excel file using Aspose.Cells C# | Aspose.Cells protect workbook after merging sheets | C# code to protect merged workbook with password | Set ProtectionType.All in Aspose.Cells | Save password‑protected Excel with Aspose.Cells .NET
// Developer Intent: Add password‑based protection to a workbook that was created by merging multiple Excel files, preventing unauthorized editing.
// Use Cases: Secure a workbook that results from merging several source files before distributing it to end users. | Integrate password protection into a web API that returns a combined Excel report. | Automate batch processing that merges workbooks and saves each as a password‑protected file.
// AI Prompts: Generate C# code that loads a merged Excel workbook, applies Workbook.Protect with a password for all protection types, and saves it as a new file using Aspose.Cells. | Show an example of handling exceptions while protecting a combined workbook with a password in Aspose.Cells for .NET. | Explain step‑by‑step how to protect the structure and windows of a merged workbook using Aspose.Cells and why ProtectionType.All is recommended.

using Aspose.Cells;
using System;
using System.IO;

// Loads a workbook created by a combine operation, applies full workbook protection (structure and windows) using Workbook.Protect with a password, and saves the result as a password‑protected Excel file.
public class ProtectMergedWorkbook
{
    public static void Run()
    {
        string inputPath = "merged.xlsx";
        string outputPath = "merged_protected.xlsx";
        string password = "MySecretPassword";

        try
        {
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that resulted from the combine operation
            Workbook mergedWorkbook = new Workbook(inputPath);

            // Apply protection to the workbook (structure and windows) with a password
            mergedWorkbook.Protect(ProtectionType.All, password);

            // Save the password‑protected workbook
            mergedWorkbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        ProtectMergedWorkbook.Run();
    }
}
