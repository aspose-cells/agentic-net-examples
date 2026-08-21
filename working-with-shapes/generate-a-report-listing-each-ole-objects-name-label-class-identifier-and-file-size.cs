// Title: List OLE objects (name, label, class ID, size) in an Excel workbook using Aspose.Cells for .NET
// Description: A C# console sample that loads an Excel file with Aspose.Cells, iterates every worksheet, reads each OleObject, and prints its shape name, display label, hexadecimal class identifier, and embedded data size in a formatted table.
// Keywords: Aspose.Cells | C# | .NET | OLE object enumeration | Excel OLE metadata | class identifier hex | object data size | worksheet OleObjects | extract OLE properties
// Common Searches: Aspose.Cells get OLE object name label | C# read OLE class identifier from Excel | list embedded OLE objects size Aspose | enumerate OleObjects worksheet Aspose.Cells | Excel OLE object report C#
// Developer Intent: Build a console program that opens an Excel workbook, extracts each OLE object's name, label, class ID (as hex), and file size, and displays the data in a table.
// Use Cases: Audit embedded OLE objects in financial models to confirm allowed content types and sizes. | Create an inventory of linked documents for compliance reporting. | Validate workbook readiness before migration by flagging oversized or unsupported OLE objects.
// AI Prompts: Generate a method that returns a collection of OLE object details (name, label, hex class ID, size) for a given Worksheet. | Extend the example to write the OLE object report to a CSV file instead of the console. | Add error handling that logs missing or corrupt OLE data while continuing the enumeration.

using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectReport
{
    // A C# console sample that loads an Excel file with Aspose.Cells, iterates every worksheet, reads each OleObject, and prints its shape name, display label, hexadecimal class identifier, and embedded data size in a formatted table.
    class Program
    {
        static void Main()
        {
            // Load the workbook that contains OLE objects (uses the provided load rule)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the collection of OLE objects for the current worksheet
                OleObjectCollection oleObjects = sheet.OleObjects;

                // If there are no OLE objects, skip to the next worksheet
                if (oleObjects.Count == 0) continue;

                Console.WriteLine($"Worksheet: {sheet.Name}");
                Console.WriteLine("------------------------------------------------------------");
                Console.WriteLine($"{"Index",5} {"Name",-20} {"Label",-20} {"Class ID",-30} {"File Size (bytes)",15}");
                Console.WriteLine(new string('-', 95));

                // Enumerate each OLE object
                for (int i = 0; i < oleObjects.Count; i++)
                {
                    OleObject ole = oleObjects[i];

                    // Name of the OLE object (shape name)
                    string name = ole.Name ?? string.Empty;

                    // Label (display label of the linked OLE object)
                    string label = ole.Label ?? string.Empty;

                    // Class identifier – convert byte[] to a hex string for readability
                    string classId = "null";
                    if (ole.ClassIdentifier != null && ole.ClassIdentifier.Length > 0)
                    {
                        StringBuilder sb = new StringBuilder(ole.ClassIdentifier.Length * 2);
                        foreach (byte b in ole.ClassIdentifier)
                            sb.AppendFormat("{0:X2}", b);
                        classId = sb.ToString();
                    }

                    // File size – length of the embedded object data (ObjectData)
                    long fileSize = ole.ObjectData?.LongLength ?? 0;

                    Console.WriteLine($"{i + 1,5} {name,-20} {label,-20} {classId,-30} {fileSize,15}");
                }

                Console.WriteLine(); // Blank line between worksheets
            }

            // Save the workbook if any modifications were made (uses the provided save rule)
            // In this example we only read data, so saving is optional.
            // workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}
