using System;
using System.IO;
using Aspose.Cells;

class BatchAddGlobalNamedRange
{
    static void Main()
    {
        // Define the list of workbook file paths to process (10 files)
        string[] workbookFiles = new string[]
        {
            @"C:\Workbooks\Book1.xlsx",
            @"C:\Workbooks\Book2.xlsx",
            @"C:\Workbooks\Book3.xlsx",
            @"C:\Workbooks\Book4.xlsx",
            @"C:\Workbooks\Book5.xlsx",
            @"C:\Workbooks\Book6.xlsx",
            @"C:\Workbooks\Book7.xlsx",
            @"C:\Workbooks\Book8.xlsx",
            @"C:\Workbooks\Book9.xlsx",
            @"C:\Workbooks\Book10.xlsx"
        };

        // Define the reference for the global named range "Quarter"
        // Adjust the sheet name and address as needed for your scenario
        const string quarterRefersTo = "=Sheet1!$A$1:$A$4";

        foreach (string filePath in workbookFiles)
        {
            // Ensure the file exists before processing
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                continue;
            }

            // Load the workbook (using the constructor that accepts a file path)
            using (Workbook workbook = new Workbook(filePath))
            {
                // Access the global (workbook‑scope) name collection
                NameCollection names = workbook.Worksheets.Names;

                // Check if a global name "Quarter" already exists
                // Global names have SheetIndex = 0
                bool quarterExists = false;
                foreach (Name existingName in names)
                {
                    if (existingName.Text.Equals("Quarter", StringComparison.OrdinalIgnoreCase) &&
                        existingName.SheetIndex == 0)
                    {
                        quarterExists = true;
                        // Update the reference in case it differs
                        existingName.RefersTo = quarterRefersTo;
                        break;
                    }
                }

                // If the name does not exist, add it as a global name
                if (!quarterExists)
                {
                    int index = names.Add("Quarter");
                    Name quarterName = names[index];
                    quarterName.RefersTo = quarterRefersTo;
                    // SheetIndex = 0 indicates a global (workbook‑scope) name
                    quarterName.SheetIndex = 0;
                }

                // Optional: sort names for better organization before saving
                workbook.Worksheets.SortNames();

                // Save the workbook back to the same file (overwrites)
                workbook.Save(filePath);
                Console.WriteLine($"Processed and saved: {filePath}");
            }
        }
    }
}