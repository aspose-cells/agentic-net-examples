// Title: Load an Excel workbook from a FileStream and write a sequential numeric series to cells A1:A10 using Aspose.Cells for .NET
// Common Searches: how to read an existing .xlsx file from a FileStream with Aspose.Cells in C# | Aspose.Cells C# example to write numbers 1 to 10 into a column range | C# code to autofill a numeric series in Excel using Aspose.Cells Workbook | saving a modified workbook to a new file after populating cells with Aspose.Cells | loop through rows and put values in Aspose.Cells cells collection C#
// Tags: load workbook from filestream Aspose.Cells | write sequential numbers to cell range Aspose.Cells | populate column A with numeric series C# | save workbook to new file Aspose.Cells | iterate rows putvalue Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample opens an existing Excel file via a FileStream, creates a Workbook object, accesses the first worksheet, writes the numbers 1 through 10 into cells A1 to A10, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from a file stream
            using (FileStream fileStream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
            {
                Workbook workbook = new Workbook(fileStream);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Fill the numeric series manually (1, 2, 3, …, 10) in cells A1:A10
                for (int row = 0; row < 10; row++)
                {
                    sheet.Cells[row, 0].PutValue(row + 1); // Column 0 corresponds to column A
                }

                // Save the modified workbook to a new file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
