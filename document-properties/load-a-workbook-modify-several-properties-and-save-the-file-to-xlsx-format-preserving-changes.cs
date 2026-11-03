// Title: Load or create an Excel workbook, rename a sheet, set default font and author, add a summary worksheet, and save as XLSX with Aspose.Cells for .NET (C#)
// Common Searches: how to load an existing Excel file or create a new workbook with Aspose.Cells C# | rename first worksheet and set default workbook font using Aspose.Cells | set the Author built‑in document property programmatically in Aspose.Cells | add a new worksheet named Summary and write a value to A1 with Aspose.Cells | save modified workbook as XLSX format using Aspose.Cells for .NET
// Tags: conditional workbook loading Aspose.Cells C# | worksheet name change Aspose.Cells | default font configuration Aspose.Cells | built‑in document property update Aspose.Cells | create and populate new worksheet Aspose.Cells | xlsx file saving Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks for an existing 'input.xlsx' file and loads it with Aspose.Cells; if the file is missing, it creates a new workbook. It renames the first worksheet, changes the workbook's default font to Calibri 11, sets the Author built‑in property to "John Doe", adds a new worksheet named "Summary" with a value in cell A1, ensures the output directory exists, and finally saves the modified workbook as 'output.xlsx' in XLSX format.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                // LoadOptions can be used to specify a password if needed.
                // Here we assume no password; adjust if required.
                var loadOptions = new LoadOptions();
                workbook = new Workbook(inputPath, loadOptions);
            }
            else
            {
                // No input file found – start with a fresh workbook
                workbook = new Workbook();
            }

            // 1. Rename the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "RenamedSheet";

            // 2. Change the default font for the entire workbook
            workbook.DefaultStyle.Font.Name = "Calibri";
            workbook.DefaultStyle.Font.Size = 11;

            // 3. Set the workbook's author property
            workbook.BuiltInDocumentProperties["Author"].Value = "John Doe";

            // 4. Add a new worksheet and write a value
            int newSheetIndex = workbook.Worksheets.Add();
            Worksheet newSheet = workbook.Worksheets[newSheetIndex];
            newSheet.Name = "Summary";
            newSheet.Cells["A1"].PutValue("Report Summary");

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
