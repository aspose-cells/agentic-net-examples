// Title: C# batch program to delete empty worksheets from multiple .xlsx workbooks using Aspose.Cells
// AI Prompts: Write C# code that iterates over every .xlsx file in a directory, loads each workbook with Aspose.Cells, removes worksheets that contain no used cells, and saves the cleaned workbook to a separate output folder. | Create a reusable C# method that determines whether an Aspose.Cells Worksheet is empty by checking MaxDataRow and MaxDataColumn, and use it inside a loop that processes a batch of Excel files.
// Common Searches: aspocells c# delete blank worksheets from a folder of excel files | how to programmatically delete empty sheets in multiple workbooks using Aspose.Cells | batch clean Excel workbooks by removing worksheets with no data in .NET | c# script to iterate over .xlsx files and drop worksheets that contain no cells | detect and delete empty worksheets in Aspose.Cells before saving
// Tags: Aspose.Cells batch worksheet removal | C# delete blank sheets .xlsx | detect empty worksheet MaxDataRow Aspose | automated Excel workbook cleanup .NET | process multiple workbooks Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;
using System.Linq;

// The program scans a specified input folder for .xlsx files, loads each workbook with Aspose.Cells, identifies worksheets that have no data (MaxDataRow/MaxDataColumn < 0), removes those empty worksheets, and saves the cleaned workbooks to an output directory.
class Program
{
    static void Main()
    {
        // Folder containing the original workbooks
        string inputFolder = @"C:\InputWorkbooks";

        // Folder where cleaned workbooks will be saved
        string outputFolder = @"C:\CleanedWorkbooks";
        Directory.CreateDirectory(outputFolder);

        // Process each .xlsx file in the input folder
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
        {
            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Identify indexes of worksheets that have no data
            var emptySheetIndexes = workbook.Worksheets
                .Cast<Worksheet>()
                .Select((sheet, idx) => new { sheet, idx })
                .Where(x => IsWorksheetEmpty(x.sheet))
                .Select(x => x.idx)
                .OrderByDescending(i => i) // delete from highest index to keep remaining indexes valid
                .ToList();

            // Remove each empty worksheet
            foreach (int idx in emptySheetIndexes)
            {
                workbook.Worksheets.RemoveAt(idx);
            }

            // Save the cleaned workbook (overwrites if same name exists in output folder)
            string fileName = Path.GetFileName(filePath);
            string outputPath = Path.Combine(outputFolder, fileName);
            workbook.Save(outputPath);
        }
    }

    // Returns true if the worksheet contains no used cells
    static bool IsWorksheetEmpty(Worksheet sheet)
    {
        // MaxDataRow/Column are -1 when the sheet has no data
        return sheet.Cells.MaxDataRow < 0 || sheet.Cells.MaxDataColumn < 0;
    }
}
