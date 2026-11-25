// Title: Classify Excel worksheets as data‑only, shape‑only, or mixed content with Aspose.Cells for .NET (C#)
// Common Searches: C# code to identify worksheets that contain only cell values using Aspose.Cells | how to check if an Excel sheet has shapes or charts with Aspose.Cells .NET | determine worksheet content type (data, shapes, mixed) in Aspose.Cells C# | using MaxDataRow and MaxDataColumn to detect data in Excel worksheets via Aspose.Cells | count shapes in a worksheet and classify sheet content with Aspose.Cells for .NET
// Tags: detect worksheet data Aspose.Cells | detect worksheet shapes Aspose.Cells | classify worksheet content Aspose.Cells | use MaxDataRow MaxDataColumn Aspose.Cells | mixed content worksheet detection Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace WorksheetContentClassifier
{
    // The example loads a workbook supplied on the command line, iterates through each worksheet, uses Worksheet.Cells.MaxDataRow and MaxDataColumn to determine if cell data exists, checks Worksheet.Shapes.Count for shapes, classifies the sheet as Data‑Only, Shape‑Only, Mixed Content, or Empty, and prints the classification along with data and shape details to the console.
    class Program
    {
        /// <param name="args">Command‑line arguments.</param>
        static void Main(string[] args)
        {
            // Validate input arguments
            if (args == null || args.Length == 0)
            {
                Console.WriteLine("Usage: WorksheetContentClassifier <path-to-workbook>");
                return;
            }

            string workbookPath = args[0];

            if (!System.IO.File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: File not found – \"{workbookPath}\"");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);
                Console.WriteLine($"Workbook \"{workbookPath}\" loaded successfully.");
                Console.WriteLine($"Number of worksheets: {workbook.Worksheets.Count}");
                Console.WriteLine();

                // Iterate through each worksheet
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    string sheetName = worksheet.Name;

                    // ----- Cell data detection -----
                    // MaxDataRow and MaxDataColumn return -1 when no cell contains data.
                    int maxDataRow = worksheet.Cells.MaxDataRow;
                    int maxDataColumn = worksheet.Cells.MaxDataColumn;
                    bool hasData = maxDataRow != -1 && maxDataColumn != -1;

                    // ----- Shape detection -----
                    // ShapeCollection inherits from CollectionBase; Count gives the number of shapes.
                    int shapeCount = worksheet.Shapes.Count;
                    bool hasShapes = shapeCount > 0;

                    // ----- Classification -----
                    string classification;
                    if (hasData && !hasShapes)
                    {
                        classification = "Data‑Only";
                    }
                    else if (!hasData && hasShapes)
                    {
                        classification = "Shape‑Only";
                    }
                    else if (hasData && hasShapes)
                    {
                        classification = "Mixed Content";
                    }
                    else
                    {
                        classification = "Empty (no data, no shapes)";
                    }

                    // Output results for the current worksheet
                    Console.WriteLine($"Worksheet: \"{sheetName}\"");
                    Console.WriteLine($"  Data present?   {(hasData ? "Yes" : "No")} (MaxDataRow={maxDataRow}, MaxDataColumn={maxDataColumn})");
                    Console.WriteLine($"  Shapes present? {(hasShapes ? "Yes" : "No")} (Count={shapeCount})");
                    Console.WriteLine($"  Classification: {classification}");
                    Console.WriteLine();
                }
            }
            catch (Exception ex)
            {
                // Capture any unexpected errors and display them
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
