// Title: Add AutoShape with adjustment guides and embed a VBA module using Aspose.Cells for .NET
// Description: Step‑by‑step guide that shows how to create a new workbook, insert a RightArrowCallout AutoShape, access its ShapeGuideCollection, add and modify guide values, and save the file as .xlsx. The example also demonstrates initializing a VBA project, adding a class module with code, and saving the workbook as an .xlsm file.
// Keywords: Aspose.Cells AutoShape guide | ShapeGuideCollection C# | adjustable shape geometry Aspose.Cells | add VBA module Aspose.Cells | RightArrowCallout shape | .NET Excel shape adjustment | save workbook as xlsm Aspose | programmatic Excel shape editing
// Common Searches: Aspose.Cells add shape guide C# | modify AutoShape adjustment values .NET | how to add VBA class module with Aspose.Cells | create RightArrowCallout shape in Excel using Aspose | save workbook with VBA project Aspose.Cells
// Developer Intent: Programmatically create an AutoShape with custom adjustment guides, update those guides, and optionally embed a VBA class module, then persist the workbook in .xlsx or .xlsm format.
// Use Cases: Generate dynamic callout diagrams where guide values control arrow length and callout position. | Adjust shape geometry on the fly to reflect data‑driven visualizations before exporting the workbook. | Add a VBA macro to a workbook produced by Aspose.Cells to provide post‑generation interactivity for end users.
// AI Prompts: Write C# code with Aspose.Cells that adds a RightArrowCallout AutoShape, creates four custom guides, changes their values, and saves the workbook as an .xlsx file. | Show how to initialize a VBA project in an Aspose.Cells workbook, add a class module named 'Helper' containing a simple Sub procedure, and save the file as an .xlsm workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Vba;

namespace AsposeCellsDeveloperGuide
{
    // Step‑by‑step guide that shows how to create a new workbook, insert a RightArrowCallout AutoShape, access its ShapeGuideCollection, add and modify guide values, and save the file as .xlsx. The example also demonstrates initializing a VBA project, adding a class module with code, and saving the workbook as an .xlsm file.
    public class ShapeGuideDeveloperGuide
    {
        public static void Run()
        {
            try
            {
                // Step 1: Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Step 2: Add an AutoShape that supports adjustment guides (e.g., RightArrowCallout)
                // Parameters: type, upperLeftRow, upperLeftColumn, upperLeftPixel, upperLeftPixel2, width, height
                Shape shape = worksheet.Shapes.AddAutoShape(
                    AutoShapeType.RightArrowCallout,
                    2,          // upperLeftRow
                    0,          // upperLeftColumn
                    2,          // upperLeftPixel
                    0,          // upperLeftPixel2
                    200,        // width
                    150);       // height

                // Step 3: Obtain the ShapeGuideCollection from the shape's geometry adjustments
                ShapeGuideCollection guideCollection = shape.Geometry.ShapeAdjustValues;

                // Step 4: Add new guides using the Add(string name, double value) method
                guideCollection.Add("adj1", 25.5);
                guideCollection.Add("adj2", 30.0);
                guideCollection.Add("adj3", 45.5);
                guideCollection.Add("adj4", 60.0);

                // Step 5: Access individual guides via the indexer and read their values
                Console.WriteLine("Initial guide values:");
                for (int i = 0; i < guideCollection.Count; i++)
                {
                    ShapeGuide guide = guideCollection[i];
                    Console.WriteLine($"Guide {i + 1} value = {guide.Value}");
                }

                // Step 6: Modify guide values as needed
                guideCollection[0].Value = 20.0;
                guideCollection[1].Value = 35.0;
                guideCollection[2].Value = 50.0;
                guideCollection[3].Value = 70.0;

                // Step 7: Verify the updated values
                Console.WriteLine("\nUpdated guide values:");
                for (int i = 0; i < guideCollection.Count; i++)
                {
                    Console.WriteLine($"Guide {i + 1} value = {guideCollection[i].Value}");
                }

                // Step 8: Save the workbook to persist the shape and its guides
                string shapeFile = "ShapeGuideDeveloperGuideDemo.xlsx";
                workbook.Save(shapeFile);
                Console.WriteLine($"\nWorkbook saved to '{Path.GetFullPath(shapeFile)}'");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in ShapeGuideDeveloperGuide: {ex.Message}");
            }
        }
    }

    // Optional: Demonstrate adding a VBA module to the same workbook
    public class VbaModuleDemo
    {
        public static void Run()
        {
            try
            {
                Workbook workbook = new Workbook();

                // Initialize VBA project (required for .xlsm format)
                VbaProject vbaProject = workbook.VbaProject;

                // Add a new class module named "Helper"
                int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Class, "Helper");

                // Retrieve the module and set its VBA code
                VbaModule vbaModule = vbaProject.Modules[moduleIndex];
                vbaModule.Codes = "Sub ShowMessage()\r\n    MsgBox \"Hello from VBA!\"\r\nEnd Sub";

                // Save the workbook with VBA project (Xlsm format)
                string vbaFile = "VbaModuleDemo.xlsm";
                workbook.Save(vbaFile, SaveFormat.Xlsm);
                Console.WriteLine($"VBA workbook saved to '{Path.GetFullPath(vbaFile)}'");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in VbaModuleDemo: {ex.Message}");
            }
        }
    }

    // Entry point for the guide
    class Program
    {
        static void Main()
        {
            ShapeGuideDeveloperGuide.Run();
            VbaModuleDemo.Run();
        }
    }
}
