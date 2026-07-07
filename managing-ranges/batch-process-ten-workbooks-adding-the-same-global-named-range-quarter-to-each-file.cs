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

        // The global named range to add
        const string globalName = "Quarter";
        // Example reference – adjust as needed
        const string refersToFormula = "=Sheet1!$A$1:$A$4";

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
                // Access the global names collection
                NameCollection names = workbook.Worksheets.Names;

                // Check if the global name already exists
                bool exists = false;
                foreach (Name n in names)
                {
                    if (n.Text.Equals(globalName, StringComparison.OrdinalIgnoreCase) && n.SheetIndex == 0)
                    {
                        exists = true;
                        break;
                    }
                }

                // Add the global named range if it does not exist
                if (!exists)
                {
                    int index = names.Add(globalName);          // Define the name
                    Name name = names[index];
                    name.RefersTo = refersToFormula;            // Set the reference
                    name.SheetIndex = 0;                        // Ensure workbook scope (0 = global)
                }

                // Optional: sort names for better organization
                workbook.Worksheets.SortNames();

                // Save the workbook back to the same file (overwrites original)
                workbook.Save(filePath);
                Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
            }
        }
    }
}