// Title: Load a workbook, list its worksheets, edit a cell, and save with Aspose.Cells for .NET
// Description: C# example that checks for an input.xlsx file, creates a placeholder workbook if missing, loads the file using Aspose.Cells, iterates the WorksheetCollection to print each sheet name, updates cell A1 on the first sheet, and saves the result as output.xlsx.
// Keywords: Aspose.Cells load workbook C# | Aspose.Cells enumerate worksheets | Aspose.Cells modify cell value | Aspose.Cells save workbook .NET | C# Excel placeholder file | WorksheetCollection Aspose.Cells | Aspose.Cells example load and edit
// Common Searches: how to load an existing Excel file with Aspose.Cells C# | list all worksheet names using Aspose.Cells .NET | update a cell after loading a workbook in Aspose.Cells | create placeholder Excel file when input is missing C# | save modified workbook with Aspose.Cells
// Developer Intent: Load an existing Excel workbook, enumerate its worksheets, change a cell value, and write the changes to a new file.
// Use Cases: Generate a dummy workbook automatically when the expected source file is absent, ensuring downstream processes do not fail. | Log or validate worksheet names by iterating the WorksheetCollection after loading a workbook. | Programmatically set or overwrite data in a specific cell (e.g., A1) of the first worksheet and persist the modification.
// AI Prompts: Write C# code that uses Aspose.Cells to open an Excel file, creates a placeholder workbook if the file does not exist, and prints each worksheet name. | Provide a snippet that changes the value of cell A1 in the first worksheet of a loaded workbook and saves the result to a different file using Aspose.Cells for .NET. | Explain how to handle missing input files gracefully when working with Aspose.Cells in a C# console application.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // C# example that checks for an input.xlsx file, creates a placeholder workbook if missing, loads the file using Aspose.Cells, iterates the WorksheetCollection to print each sheet name, updates cell A1 on the first sheet, and saves the result as output.xlsx.
    public class LoadWorkbookAndAccessWorksheets
    {
        public static void Run()
        {
            // Path to the existing Excel file
            string dataDir = "YourDocumentDirectory/";
            string inputFile = Path.Combine(dataDir, "input.xlsx");

            try
            {
                // Ensure the input file exists; create a placeholder if missing
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    Workbook placeholder = new Workbook();
                    placeholder.Save(inputFile);
                    Console.WriteLine($"Created placeholder workbook at {inputFile}");
                }

                // Load the workbook from the file
                Workbook workbook = new Workbook(inputFile);

                // Access the worksheet collection
                WorksheetCollection worksheets = workbook.Worksheets;

                // Iterate through worksheets and display their names
                foreach (Worksheet sheet in worksheets)
                {
                    Console.WriteLine("Worksheet: " + sheet.Name);
                }

                // Manipulate the first worksheet: set a value in cell A1
                Worksheet firstSheet = worksheets[0];
                firstSheet.Cells["A1"].PutValue("Loaded and Modified");

                // Save the modified workbook
                string outputFile = Path.Combine(dataDir, "output.xlsx");
                workbook.Save(outputFile);
                Console.WriteLine("Workbook saved to " + outputFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            LoadWorkbookAndAccessWorksheets.Run();
        }
    }
}
