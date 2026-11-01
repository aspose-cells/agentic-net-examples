// Title: Copy a specific cell range to a new workbook and set the workbook's creation date using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# copy range A1:C10 to a new workbook | How to set the CreatedTime property of an Excel file with Aspose.Cells | Copy cells from one worksheet to another workbook while preserving styles using Aspose.Cells | Set workbook metadata creation date to the current timestamp in C# Aspose.Cells | Example of copying a range and updating workbook properties with Aspose.Cells .NET
// Tags: copy range A1:C10 into new workbook Aspose.Cells | assign current timestamp to CreatedTime Aspose.Cells | preserve cell formatting during range copy Aspose.Cells | C# Aspose.Cells workbook metadata example | copy cells between workbooks Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample loads source.xlsx, copies the cells A1:C10 (including values and styles) from the first worksheet into a newly created workbook's first sheet named "CopiedData", sets the destination workbook's CreatedTime built‑in document property to DateTime.Now, and saves the result as destination.xlsx while handling missing files and copy errors.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string destinationPath = "destination.xlsx";

            // Verify source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Define the range to copy (e.g., A1:C10)
            CellArea sourceRange = CellArea.CreateCellArea("A1", "C10");

            // Create a new workbook (destination)
            Workbook destinationWorkbook = new Workbook();

            // Get the first worksheet in the destination workbook and rename it
            Worksheet destSheet = destinationWorkbook.Worksheets[0];
            destSheet.Name = "CopiedData";

            // Calculate rows and columns count for the copy operation
            int totalRows = sourceRange.EndRow - sourceRange.StartRow + 1;
            int totalColumns = sourceRange.EndColumn - sourceRange.StartColumn + 1;

            // Copy the defined range from the source workbook to the destination worksheet starting at A1
            try
            {
                Worksheet srcSheet = sourceWorkbook.Worksheets[0];
                for (int i = 0; i < totalRows; i++)
                {
                    for (int j = 0; j < totalColumns; j++)
                    {
                        // Source cell
                        Cell srcCell = srcSheet.Cells[sourceRange.StartRow + i, sourceRange.StartColumn + j];
                        // Destination cell
                        Cell destCell = destSheet.Cells[i, j];

                        // Copy value
                        destCell.PutValue(srcCell.Value);
                        // Copy style
                        destCell.SetStyle(srcCell.GetStyle());
                    }
                }
            }
            catch (Exception copyEx)
            {
                Console.WriteLine($"Error during copy operation: {copyEx.Message}");
                return;
            }

            // Set the workbook's creation date metadata to the current timestamp
            destinationWorkbook.BuiltInDocumentProperties.CreatedTime = DateTime.Now;

            // Save the new workbook
            destinationWorkbook.Save(destinationPath);
            Console.WriteLine($"Workbook saved successfully to {destinationPath}");
        }
        catch (Exception ex)
        {
            // Log or display any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
