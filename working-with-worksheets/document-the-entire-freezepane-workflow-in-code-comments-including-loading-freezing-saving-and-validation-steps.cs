// Title: How to Load or Create an Excel Workbook, Apply Freeze Panes to the First Row and Column, Save, and Verify the Output Using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to open an existing Excel file (or create a new workbook if the file is missing), freeze the first row and first column, save the workbook to a given path, and return a boolean indicating whether the save succeeded. | Show how to adjust the FreezePanes method parameters in Aspose.Cells to freeze any custom range of rows and columns, then save the workbook and log the result of the operation. | Provide a C# snippet that wraps the freeze‑pane workflow in a try‑catch block, logs detailed error information, and prints a success or failure message after attempting to save the workbook.
// Common Searches: asp.net c# apply freeze panes to first row and column with Aspose.Cells | Aspose.Cells create workbook if file not found then freeze panes and save | C# verify Excel file was saved after applying freeze panes using Aspose.Cells | example code for FreezePanes method with custom row and column indices in Aspose.Cells .NET
// Tags: Aspose.Cells freeze panes C# example | load or create workbook Aspose.Cells .NET | apply FreezePanes method Excel C# | save workbook and check file existence Aspose.Cells | Excel freeze pane validation C#

using Aspose.Cells;
using System;
using System.IO;

// The sample program checks for an "input.xlsx" file, loads it if present or creates a new workbook with sample data otherwise, accesses the first worksheet, freezes the first row and column using FreezePanes(1,1,1,1), saves the result as "output.xlsx", and then verifies that the output file exists, printing a success or failure message.
class FreezePaneWorkflow
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook or create a new one if the input file is missing.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a new workbook with one worksheet.
                workbook.Worksheets[0].Cells["A1"].PutValue("Sample Data");
            }

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Apply freeze pane: freeze the first row and first column.
            sheet.FreezePanes(1, 1, 1, 1);

            // Save the workbook with the freeze pane applied.
            workbook.Save(outputPath);

            // Simple validation: ensure the output file was created.
            if (File.Exists(outputPath))
            {
                Console.WriteLine("Freeze pane applied and file saved successfully.");
            }
            else
            {
                Console.WriteLine("Failed to save the output file.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
