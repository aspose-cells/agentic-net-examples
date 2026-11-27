// Title: Create a formula that links to an external workbook with a relative path and auto‑updates when the source moves using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set formula to external workbook using relative path | how to make external Excel link update after moving source file with Aspose.Cells | create external reference formula in .NET workbook that handles missing source file | C# Aspose.Cells example linking cell to another workbook Sheet1!B2 | Aspose.Cells warning for absent external file when creating link
// Tags: external workbook formula Aspose.Cells | relative path reference Excel Aspose.Cells | auto‑refresh external link Aspose.Cells | missing external file handling Aspose.Cells | C# create workbook with external cell reference

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, renames the first worksheet to "Summary", defines a relative path to an external file "ExternalData.xlsx", warns if the file is missing, sets cell A1 formula to reference B2 on Sheet1 of the external workbook, and saves the result as "ResultWithExternalLink.xlsx".
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and give it a meaningful name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Summary";

            // Define the external workbook file name (use a relative path if possible)
            string externalFile = "ExternalData.xlsx";

            // Warn if the external file does not exist; the link will still be created
            if (!File.Exists(externalFile))
            {
                Console.WriteLine($"Warning: External file '{externalFile}' not found. The link may be unresolved until the file exists.");
            }

            // Set a formula that references cell B2 in Sheet1 of the external workbook
            sheet.Cells["A1"].Formula = $"='[{externalFile}]Sheet1'!B2";

            // Save the workbook containing the external reference
            string resultFile = "ResultWithExternalLink.xlsx";
            workbook.Save(resultFile);
            Console.WriteLine($"Workbook saved successfully to '{resultFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
