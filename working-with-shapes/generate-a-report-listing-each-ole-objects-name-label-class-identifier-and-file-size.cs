// Title: Create an Excel report of OLE object names, labels, class identifiers, and file sizes using Aspose.Cells for .NET
// AI Prompts: Write a C# console application that loads a source workbook with Aspose.Cells, loops through each worksheet's OleObjectCollection, extracts the OLE object's Name, Label, ClassIdentifier (converted to a hex string), and the length of its ObjectData, writes these values to a new worksheet, auto‑fits the columns, and saves the report workbook. | Include robust error handling that checks for the existence of the input file, gracefully handles OLE objects without embedded data, and prints a confirmation message with the output file path.
// Common Searches: Aspose.Cells C# list all OLE objects in an Excel workbook and export their properties | How to retrieve the ClassIdentifier of embedded OLE objects using Aspose.Cells for .NET | Create a summary sheet of OLE object metadata (name, label, size) with Aspose.Cells | Get embedded file size of OLE objects in a .xlsx file via Aspose.Cells API | Generate a report of OLE objects across worksheets using Aspose.Cells C# example
// Tags: Aspose.Cells OLE object enumeration to Excel | extract OLE metadata using Aspose.Cells C# | convert ClassIdentifier byte array to hex Aspose.Cells | calculate embedded OLE file size Aspose.Cells | auto-fit columns Aspose.Cells report

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectReport
{
    // The program loads an input workbook, iterates every worksheet's OleObjectCollection, extracts each OLE object's Name, Label, ClassIdentifier (as a hex string), and embedded file size, writes these details to a new worksheet, auto‑fits columns, and saves the generated report.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the source workbook that contains OLE objects
                Workbook sourceWorkbook = new Workbook(inputPath);

                // Create a new workbook for the report
                Workbook reportWorkbook = new Workbook();
                Worksheet reportSheet = reportWorkbook.Worksheets[0];
                reportSheet.Name = "OLE Report";

                // Write header row
                reportSheet.Cells[0, 0].PutValue("Worksheet");
                reportSheet.Cells[0, 1].PutValue("OLE Name");
                reportSheet.Cells[0, 2].PutValue("Label");
                reportSheet.Cells[0, 3].PutValue("Class Identifier");
                reportSheet.Cells[0, 4].PutValue("File Size (bytes)");

                int reportRow = 1; // start after header

                // Iterate through each worksheet in the source workbook
                foreach (Worksheet ws in sourceWorkbook.Worksheets)
                {
                    // Access the collection of OLE objects on the current worksheet
                    OleObjectCollection oleObjects = ws.OleObjects;

                    // Loop through each OLE object
                    foreach (OleObject ole in oleObjects)
                    {
                        // Retrieve required properties
                        string oleName = ole.Name;                     // OLE object's name
                        string oleLabel = ole.Label;                   // OLE object's label (display text)

                        // ClassIdentifier returns a byte[]; convert to a readable hex string
                        byte[] classIdBytes = ole.ClassIdentifier;
                        string classId = classIdBytes != null && classIdBytes.Length > 0
                            ? BitConverter.ToString(classIdBytes).Replace("-", "")
                            : string.Empty;

                        // Get embedded file size if available
                        int fileSize = 0;
                        if (ole.ObjectData != null)
                        {
                            fileSize = ole.ObjectData.Length;
                        }

                        // Write data to the report sheet
                        reportSheet.Cells[reportRow, 0].PutValue(ws.Name);
                        reportSheet.Cells[reportRow, 1].PutValue(oleName);
                        reportSheet.Cells[reportRow, 2].PutValue(oleLabel);
                        reportSheet.Cells[reportRow, 3].PutValue(classId);
                        reportSheet.Cells[reportRow, 4].PutValue(fileSize);

                        reportRow++;
                    }
                }

                // Auto-fit columns for better readability
                reportSheet.AutoFitColumns();

                // Save the report workbook
                reportWorkbook.Save(outputPath);
                Console.WriteLine($"Report saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
