// Title: How to retrieve positions, sizes, and placement of all TextBox shapes in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# list all TextBox shapes on a worksheet and get their row and column indices | How to read the width and height of Excel TextBox objects using Aspose.Cells .NET | C# code to output TextBox placement and coordinates from an .xlsx file with Aspose.Cells | Retrieve TextBox collection and print upper‑left cell indices in Aspose.Cells for .NET | Extract TextBox properties such as text, size, and placement from a workbook using Aspose.Cells
// Tags: Aspose.Cells textbox position extraction | Aspose.Cells textbox size retrieval | C# enumerate worksheet TextBox collection | Aspose.Cells get textbox placement property | Aspose.Cells retrieve textbox row column indices

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace TextBoxPositionExtractor
{
    // The example loads an Excel workbook with Aspose.Cells, accesses the first worksheet, enumerates its TextBoxCollection, and prints each TextBox's text, upper‑left row and column, width, height (pixels), and placement setting.
    class Program
    {
        static void Main(string[] args)
        {
            // Determine the input workbook path.
            string workbookPath;
            if (args.Length > 0 && File.Exists(args[0]))
            {
                workbookPath = args[0];
            }
            else
            {
                // Default file name in the executable directory.
                workbookPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "input.xlsx");
                if (!File.Exists(workbookPath))
                {
                    Console.Error.WriteLine("Error: No input workbook found. Provide a valid path as a command‑line argument or place 'input.xlsx' beside the executable.");
                    return;
                }
            }

            try
            {
                // Load the workbook.
                Workbook workbook = new Workbook(workbookPath);

                // Choose the first worksheet (modify if a specific sheet is required).
                Worksheet worksheet = workbook.Worksheets[0];
                Console.WriteLine($"Processing worksheet: \"{worksheet.Name}\"");

                // Retrieve the collection of TextBox objects.
                TextBoxCollection textBoxes = worksheet.TextBoxes;

                if (textBoxes.Count == 0)
                {
                    Console.WriteLine("No TextBox controls were found on this worksheet.");
                    return;
                }

                Console.WriteLine($"Found {textBoxes.Count} TextBox(es).");

                // Iterate through each TextBox and output its position and size.
                for (int i = 0; i < textBoxes.Count; i++)
                {
                    TextBox tb = textBoxes[i];

                    // Position properties (row/column indices).
                    int upperRow = tb.UpperLeftRow;
                    int upperColumn = tb.UpperLeftColumn;

                    // Size properties (in pixels).
                    int heightPixels = tb.Height;
                    int widthPixels = tb.Width;

                    Console.WriteLine($"TextBox #{i + 1}:");
                    Console.WriteLine($"  Text               : \"{tb.Text}\"");
                    Console.WriteLine($"  UpperLeftRow       : {upperRow}");
                    Console.WriteLine($"  UpperLeftColumn    : {upperColumn}");
                    Console.WriteLine($"  Width (pixels)     : {widthPixels}");
                    Console.WriteLine($"  Height (pixels)    : {heightPixels}");
                    Console.WriteLine($"  Placement          : {tb.Placement}");
                    Console.WriteLine();
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An exception occurred while processing the workbook: {ex.Message}");
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
