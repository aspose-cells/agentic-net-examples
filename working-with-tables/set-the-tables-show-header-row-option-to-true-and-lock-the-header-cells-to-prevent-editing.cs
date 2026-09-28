// Title: Enable the header row and lock its cells for the first Excel table using Aspose.Cells in C#
// AI Prompts: Write C# code that sets ListObject.ShowHeaderRow to true, applies a locked style to the header cells, and protects the worksheet with Aspose.Cells. | Generate a snippet that creates a Style with IsLocked = true, applies it to the header range of a table, and saves the workbook. | Provide a step‑by‑step example of loading an existing .xlsx, enabling the table header row, locking the header row, and protecting the sheet using Aspose.Cells. | Show how to protect only the header row while leaving other cells editable in an Aspose.Cells workbook.
// Common Searches: aspacells c# enable table header row and lock it | how to lock the header row of an Excel table using Aspose.Cells .NET | set ShowHeaderRow true and protect the worksheet with Aspose.Cells example | apply a locked cell format to the header of a table in C# Aspose.Cells | protect an Excel sheet after locking the table header using Aspose.Cells | c# Aspose.Cells lock header row of the first ListObject
// Tags: Aspose.Cells ListObject ShowHeaderRow true | Aspose.Cells header row cell lock | Aspose.Cells worksheet protection with locked header | Aspose.Cells create locked cell style C# | Aspose.Cells first table header protection .NET

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

// Loads an existing workbook, enables the first table's header row, applies a locked style to the header cells, protects the worksheet, and saves the modified file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Process the first table if present
            if (sheet.ListObjects.Count > 0)
            {
                ListObject table = sheet.ListObjects[0];
                table.ShowHeaderRow = true;

                // Header row position and size
                int headerRow = table.StartRow;               // first row of the table (header)
                int startColumn = table.StartColumn;          // first column of the table
                int columnCount = table.ListColumns.Count;    // number of columns in the table

                // Create a style that locks cells
                Style lockedStyle = workbook.CreateStyle();
                lockedStyle.IsLocked = true;

                // Specify that only the Locked flag should be applied
                StyleFlag styleFlag = new StyleFlag();
                styleFlag.Locked = true;

                // Apply the locked style to the header row cells
                AsposeRange headerRange = sheet.Cells.CreateRange(headerRow, startColumn, 1, columnCount);
                headerRange.ApplyStyle(lockedStyle, styleFlag);
            }

            // Protect the worksheet so locked cells cannot be edited
            sheet.Protect(ProtectionType.All);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
