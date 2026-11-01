// Title: How to set a default column width of 15 characters for every column in an Aspose.Cells worksheet using C#
// AI Prompts: Use Aspose.Cells for .NET to iterate over each column index and apply a column width of 15 characters. | Generate a workbook, retrieve its primary sheet, and assign a 15‑character width to each column via the Cells API.
// Common Searches: C# Aspose.Cells set column width for all columns in a worksheet | default column width 15 characters Aspose.Cells .NET example | loop through columns to apply uniform width using Aspose.Cells | how to set column width for entire sheet in Aspose.Cells C# | Aspose.Cells bulk column width setting for Excel file
// Tags: Aspose.Cells SetColumnWidth all columns | C# set default column width Excel worksheet | loop column width Aspose.Cells workbook | initialize worksheet with fixed column width | save workbook after column width adjustment Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Creates a new Workbook, accesses the first Worksheet, and iterates over the 256 columns, calling Cells.SetColumnWidth(col, 15) to assign a 15‑character width to each column before saving the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet in the workbook
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the width of all columns to 15 characters
            const double columnWidth = 15;
            const int totalColumns = 256; // Excel supports up to 256 columns in older formats

            // Aspose.Cells does not provide a bulk SetColumnWidth overload in all versions,
            // so set the width column by column.
            for (int col = 0; col < totalColumns; col++)
            {
                worksheet.Cells.SetColumnWidth(col, columnWidth);
            }

            // Define output file path
            string outputPath = "output.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (string.IsNullOrEmpty(outputDir))
            {
                outputDir = Directory.GetCurrentDirectory();
            }

            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
