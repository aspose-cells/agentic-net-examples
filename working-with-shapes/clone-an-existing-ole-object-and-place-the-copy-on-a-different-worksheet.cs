// Title: How to clone an OLE object from one worksheet to another using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# copy OLE object to a different worksheet | duplicate embedded OLE object in Excel with Aspose.Cells .NET | preserve OLE object properties while cloning using Aspose.Cells API | programmatically add OLE object to a new worksheet in C# | move OLE objects between worksheets in an Excel file with Aspose.Cells
// Tags: clone OLE object Aspose.Cells | copy OLE object between worksheets .NET | preserve OLE object properties Aspose.Cells | add OLE object to new worksheet C# | Aspose.Cells embedded OLE handling

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectCloneDemo
{
    // The example loads an Excel workbook, extracts the first OLE object's image and data, creates a new worksheet, adds a new OLE object at a specified cell using the extracted image, copies all original properties (display as icon, label, ProgID, format, link settings, visual attributes, etc.), and saves the workbook as a new file.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input and output workbooks
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input workbook \"{inputPath}\" not found.");
                Console.WriteLine("Please place an Excel file containing at least one OLE object at the specified location and retry.");
                return;
            }

            try
            {
                // Load the source workbook
                Workbook workbook = new Workbook(inputPath);
                Console.WriteLine($"Loaded workbook \"{inputPath}\" successfully.");

                // Choose the source worksheet (first worksheet by default)
                Worksheet sourceSheet = workbook.Worksheets[0];
                Console.WriteLine($"Source worksheet: \"{sourceSheet.Name}\".");

                // Get the collection of OLE objects on the source worksheet
                OleObjectCollection sourceOleCollection = sourceSheet.OleObjects;

                if (sourceOleCollection.Count == 0)
                {
                    Console.WriteLine("The source worksheet does not contain any OLE objects to clone.");
                    return;
                }

                // For demonstration, clone the first OLE object in the collection
                OleObject sourceOle = sourceOleCollection[0];
                Console.WriteLine($"Cloning OLE object from source worksheet.");

                // Retrieve essential data from the source OLE object
                byte[] imageData = sourceOle.ImageData;          // Icon or preview image
                byte[] objectData = sourceOle.ObjectData;        // Embedded file data
                bool displayAsIcon = sourceOle.DisplayAsIcon;
                string label = sourceOle.Label;
                string progId = sourceOle.ProgID;
                FileFormatType fileFormat = sourceOle.FileFormatType;
                bool isLink = sourceOle.IsLink;
                string objectSourceFullName = sourceOle.ObjectSourceFullName;
                string imageSourceFullName = sourceOle.ImageSourceFullName;

                // Create (or get) the target worksheet where the clone will be placed
                Worksheet targetSheet = workbook.Worksheets.Add("ClonedOLE");
                Console.WriteLine($"Created target worksheet: \"{targetSheet.Name}\".");

                // Determine placement for the cloned OLE object
                int targetRow = 5;      // Upper left row (zero‑based)
                int targetColumn = 5;   // Upper left column (zero‑based)
                int targetWidth = sourceOle.Width;
                int targetHeight = sourceOle.Height;

                // Add a new OLE object to the target worksheet using the image data
                int clonedOleIndex = targetSheet.OleObjects.Add(targetRow, targetColumn, targetWidth, targetHeight, imageData);
                OleObject clonedOle = targetSheet.OleObjects[clonedOleIndex];

                // Populate the cloned OLE object with the original data and properties
                clonedOle.ObjectData = objectData;
                clonedOle.DisplayAsIcon = displayAsIcon;
                clonedOle.Label = label;
                clonedOle.ProgID = progId;
                clonedOle.FileFormatType = fileFormat;
                clonedOle.IsLink = isLink;
                clonedOle.ObjectSourceFullName = objectSourceFullName;
                clonedOle.ImageSourceFullName = imageSourceFullName;

                // Optionally copy additional visual properties
                clonedOle.IsAutoSize = sourceOle.IsAutoSize;
                clonedOle.AutoLoad = sourceOle.AutoLoad;
                clonedOle.AutoUpdate = sourceOle.AutoUpdate;
                clonedOle.ZOrderPosition = sourceOle.ZOrderPosition;
                clonedOle.RotationAngle = sourceOle.RotationAngle;
                clonedOle.IsAspectRatioLocked = sourceOle.IsAspectRatioLocked;

                Console.WriteLine($"Cloned OLE object added to worksheet \"{targetSheet.Name}\" at row {targetRow}, column {targetColumn}.");

                // Save the modified workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully as \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
