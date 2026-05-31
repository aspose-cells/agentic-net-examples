using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class OleObjectAssignNewGuidDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add an OLE object with placeholder data
                int oleIndex = sheet.OleObjects.Add(10, 10, 200, 100, new byte[0]);
                OleObject oleObject = sheet.OleObjects[oleIndex];

                // Assign a new GUID to ClassIdentifier
                Guid newGuid = Guid.NewGuid();
                oleObject.ClassIdentifier = newGuid.ToByteArray();

                // Optionally set ProgID
                oleObject.ProgID = "Excel.Sheet.12";

                // Save the workbook
                string outputPath = "OleObjectWithNewGuid.xlsx";
                workbook.Save(outputPath);

                // Verify that the GUID was persisted
                if (File.Exists(outputPath))
                {
                    Workbook loadedWorkbook = new Workbook(outputPath);
                    OleObject loadedOle = loadedWorkbook.Worksheets[0].OleObjects[0];
                    Guid storedGuid = new Guid(loadedOle.ClassIdentifier);
                    Console.WriteLine("Stored ClassIdentifier GUID: " + storedGuid);
                }
                else
                {
                    Console.WriteLine($"File not found: {outputPath}");
                }
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
            OleObjectAssignNewGuidDemo.Run();
        }
    }
}