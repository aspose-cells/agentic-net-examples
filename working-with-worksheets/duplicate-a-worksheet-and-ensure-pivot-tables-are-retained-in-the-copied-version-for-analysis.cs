using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPivotCopyDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string sourceFile = "source.xlsx";
                const string outputFile = "output.xlsx";
                const string sourceSheetName = "DataSheet";

                // Verify source file exists
                if (!File.Exists(sourceFile))
                {
                    Console.WriteLine($"Source file \"{sourceFile}\" not found.");
                    return;
                }

                // Load workbook
                Workbook workbook = new Workbook(sourceFile);

                // Ensure the source worksheet exists
                Worksheet sourceSheet = workbook.Worksheets[sourceSheetName];
                if (sourceSheet == null)
                {
                    Console.WriteLine($"Worksheet \"{sourceSheetName}\" does not exist in the workbook.");
                    return;
                }

                // Duplicate the worksheet (copies data, formats, and pivot tables)
                int copiedIndex = workbook.Worksheets.AddCopy(sourceSheetName);
                Worksheet copiedSheet = workbook.Worksheets[copiedIndex];

                // Rename the copied worksheet, ensuring the name is unique
                string newName = sourceSheetName + "_Copy";
                int suffix = 1;
                while (workbook.Worksheets[newName] != null)
                {
                    newName = $"{sourceSheetName}_Copy{suffix++}";
                }
                copiedSheet.Name = newName;

                // Refresh pivot tables in the copied sheet
                copiedSheet.RefreshPivotTables();

                // Save the modified workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved as \"{outputFile}\".");
            }
            catch (CellsException ex)
            {
                Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}