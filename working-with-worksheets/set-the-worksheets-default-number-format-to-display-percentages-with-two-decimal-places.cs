// Title: How to set a worksheet’s default number format to a two‑decimal‑place percentage using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set default cell format to percentage with two decimal places | apply a global 0.00% number format to a new workbook in Aspose.Cells | Aspose.Cells .NET change workbook DefaultStyle to percentage format | set default number format for all cells in Aspose.Cells C# example | how to use custom number format 0.00% with Aspose.Cells workbook
// Tags: Aspose.Cells defaultstyle percentage | workbook-wide number format Aspose.Cells | custom percentage style .NET | set default worksheet format C# | global numeric format Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, modifies its DefaultStyle to use the custom number format "0.00%" (percentage with two decimal places), ensures the output directory exists, saves the file as output.xlsx, and confirms successful creation.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Set the default number format to percentage with two decimal places
            Style defaultStyle = workbook.DefaultStyle;
            defaultStyle.Custom = "0.00%";
            workbook.DefaultStyle = defaultStyle; // apply the modified style

            // Ensure the output directory exists
            string outputPath = "output.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
