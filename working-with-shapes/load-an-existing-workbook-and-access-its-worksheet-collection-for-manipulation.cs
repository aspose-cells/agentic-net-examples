// Title: Load an Excel workbook, rename the first worksheet, add a new sheet, and save the file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to open an Excel file if it exists, otherwise instantiate a new workbook, rename the first worksheet, add a worksheet named "NewSheet", and save the result to a different file. | Create a C# routine that checks for an existing .xlsx file, loads it with Aspose.Cells, modifies the Worksheets collection (change the name of the first sheet and insert an extra sheet), then persists the workbook to a new location.
// Common Searches: asp.net load existing workbook and rename first sheet with Aspose.Cells | c# add a new worksheet to an Excel file using Aspose.Cells | how to check file existence before opening workbook in Aspose.Cells C# | save modified Excel workbook to a new file path with Aspose.Cells
// Tags: load workbook from file Aspose.Cells | change worksheet name Aspose.Cells | insert additional worksheet Aspose.Cells | conditional workbook creation Aspose.Cells | export modified workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates how to load an existing Excel workbook (or create a new one if the file is missing), access its Worksheets collection, rename the first worksheet, add a new worksheet, and save the updated workbook using Aspose.Cells for .NET in C#.
class WorkbookExample
{
    static void Main()
    {
        try
        {
            // Specify the path to the existing Excel file
            string inputPath = "input.xlsx";
            Workbook workbook;

            // Load the workbook if the file exists; otherwise create a new workbook
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one sheet
            }

            // Access the collection of worksheets in the workbook
            WorksheetCollection sheets = workbook.Worksheets;

            // Example manipulation: rename the first worksheet
            if (sheets.Count > 0)
            {
                Worksheet firstSheet = sheets[0];
                firstSheet.Name = "RenamedSheet";
            }

            // Example manipulation: add a new worksheet
            sheets.Add("NewSheet");

            // Save the modified workbook (optional)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
