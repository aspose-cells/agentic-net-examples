// Title: Check that a converted Excel workbook contains no ListObject tables using Aspose.Cells for .NET
// AI Prompts: Write a C# console program that loads an Excel file with Aspose.Cells and returns true if any worksheet has ListObject tables. | Enhance the validation tool to output the names of worksheets that still contain ListObjects after conversion. | Create a reusable method `bool HasListObjects(string workbookPath)` that uses Aspose.Cells to detect ListObjects in a workbook.
// Common Searches: how to programmatically verify that an Excel file has no tables after conversion using Aspose.Cells C# | C# Aspose.Cells detect remaining ListObject objects in a workbook | sample code to iterate worksheets and check ListObjects collection Aspose.Cells .NET | validate removal of Excel tables (ListObjects) after saving with Aspose.Cells
// Tags: Aspose.Cells ListObject enumeration .NET | validate workbook tables removal Aspose.Cells | detect Excel ListObjects using C# | check worksheet ListObjects Aspose.Cells | C# workbook validation for table absence

using System;
using Aspose.Cells;

// Loads a workbook with Aspose.Cells, iterates through each worksheet to examine the ListObjects collection, and prints a pass/fail message indicating whether any ListObject tables remain in the file.
class WorkbookListObjectValidator
{
    static void Main(string[] args)
    {
        // Expect the path to the converted workbook as the first argument.
        if (args.Length == 0)
        {
            Console.WriteLine("Please provide the path to the workbook file.");
            return;
        }

        string workbookPath = args[0];

        // Load the workbook (using the provided load rule).
        Workbook workbook = new Workbook(workbookPath);

        bool containsListObjects = false;

        // Iterate through all worksheets and check for ListObjects.
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            if (sheet.ListObjects.Count > 0)
            {
                containsListObjects = true;
                break;
            }
        }

        // Output validation result.
        if (containsListObjects)
        {
            Console.WriteLine("Validation failed: The workbook still contains ListObject(s).");
        }
        else
        {
            Console.WriteLine("Validation passed: No ListObject objects found in the workbook.");
        }
    }
}
