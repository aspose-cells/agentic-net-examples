// Title: Validate Unique Worksheet TabId and UniqueId After Using AddCopy in Aspose.Cells for .NET
// Description: C# example that creates a workbook, adds two worksheets, copies them with AddCopy (by index and by name), and then scans all worksheets to confirm each TabId and UniqueId is distinct, logging any duplicates before saving the file.
// Keywords: Aspose.Cells | .NET | C# worksheet TabId | UniqueId validation | AddCopy duplicate ID | Excel sheet identifier | worksheet copy integrity | Aspose.Cells example
// Common Searches: Aspose.Cells unique TabId after AddCopy | C# check worksheet UniqueId uniqueness | detect duplicate worksheet IDs in Aspose.Cells | AddCopy sheet identifier conflict | ensure distinct worksheet IDs in generated Excel
// Developer Intent: Confirm that every worksheet retains a unique TabId (and UniqueId) after copy operations.
// Use Cases: Run validation after copying sheets to prevent ID collisions in reporting tools. | Embed the check in automated unit tests for Excel generation pipelines. | Log duplicate identifiers during workbook creation to aid debugging. | Integrate the logic into CI/CD workflows that produce Excel files. | Verify identifier integrity before publishing workbooks to end users.
// AI Prompts: Generate a C# method that returns a list of worksheet names with duplicate TabId values in an Aspose.Cells workbook. | Create an NUnit test that asserts no duplicate UniqueId exists after calling AddCopy on multiple worksheets. | Write code that throws an InvalidOperationException when a duplicate TabId is detected during workbook processing. | Provide a PowerShell script that uses Aspose.Cells to validate worksheet IDs in an existing .xlsx file. | Suggest how to extend the validation to include custom metadata checks for each worksheet.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace WorksheetIdValidationDemo
{
    // C# example that creates a workbook, adds two worksheets, copies them with AddCopy (by index and by name), and then scans all worksheets to confirm each TabId and UniqueId is distinct, logging any duplicates before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add initial worksheets and put some sample data
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Original1";
            sheet1.Cells["A1"].PutValue("Data in Original1");

            Worksheet sheet2 = workbook.Worksheets.Add("Original2");
            sheet2.Cells["A1"].PutValue("Data in Original2");

            // Perform copy operations using AddCopy (by index and by name)
            int copyIndex1 = workbook.Worksheets.AddCopy(0); // copy of Original1
            Worksheet copiedSheet1 = workbook.Worksheets[copyIndex1];
            copiedSheet1.Name = "CopyOfOriginal1";

            int copyIndex2 = workbook.Worksheets.AddCopy("Original2"); // copy of Original2
            Worksheet copiedSheet2 = workbook.Worksheets[copyIndex2];
            copiedSheet2.Name = "CopyOfOriginal2";

            // Validate that each worksheet has a unique TabId (internal sheet identifier)
            HashSet<int> tabIds = new HashSet<int>();
            bool duplicateFound = false;

            foreach (Worksheet ws in workbook.Worksheets)
            {
                int id = ws.TabId;
                if (!tabIds.Add(id))
                {
                    // Duplicate TabId detected
                    duplicateFound = true;
                    Console.WriteLine($"Duplicate TabId found on worksheet '{ws.Name}' with TabId = {id}");
                }
            }

            if (!duplicateFound)
            {
                Console.WriteLine("All worksheets have unique TabId values.");
            }

            // Optionally, also verify UniqueId uniqueness
            HashSet<string> uniqueIds = new HashSet<string>();
            foreach (Worksheet ws in workbook.Worksheets)
            {
                string uid = ws.UniqueId;
                if (!uniqueIds.Add(uid))
                {
                    Console.WriteLine($"Duplicate UniqueId found on worksheet '{ws.Name}' with UniqueId = {uid}");
                }
            }

            // Save the workbook (demonstrates lifecycle usage)
            workbook.Save("WorksheetIdValidationResult.xlsx");
        }
    }
}
