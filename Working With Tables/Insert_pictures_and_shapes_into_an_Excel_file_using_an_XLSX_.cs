using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class InsertPicturesAndShapes
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a picture using a stream (if the file exists)
        const string pic1Path = "sample.jpg";
        if (File.Exists(pic1Path))
        {
            using (FileStream fs = new FileStream(pic1Path, FileMode.Open, FileAccess.Read))
            {
                // Places the picture from row 1, column 1 to row 5, column 5
                sheet.Pictures.Add(1, 1, 5, 5, fs);
            }
        }

        // Add a picture using a stream
        const string pic2Path = "sample2.png";
        if (File.Exists(pic2Path))
        {
            using (FileStream fs = new FileStream(pic2Path, FileMode.Open, FileAccess.Read))
            {
                // Places the picture with its upper‑left corner at row 6, column 1
                sheet.Pictures.Add(6, 1, fs);
            }
        }

        // Add a picture using ShapeCollection.AddPicture with a stream
        const string pic3Path = "sample3.bmp";
        ShapeCollection shapes = sheet.Shapes;
        if (File.Exists(pic3Path))
        {
            using (FileStream fs = new FileStream(pic3Path, FileMode.Open, FileAccess.Read))
            {
                // Places the picture from row 10, column 2 to row 12, column 4
                shapes.AddPicture(10, 2, 12, 4, fs);
            }
        }

        // Add an SVG image using ShapeCollection.AddSvg
        const string svgPath = "vector.svg";
        if (File.Exists(svgPath))
        {
            byte[] svgData;
            using (FileStream fs = new FileStream(svgPath, FileMode.Open, FileAccess.Read))
            {
                svgData = new byte[fs.Length];
                fs.Read(svgData, 0, svgData.Length);
            }
            // topRow = 15, top offset = 0, leftColumn = 3, left offset = 0, height = -1, width = -1 (auto size)
            shapes.AddSvg(15, 0, 3, 0, -1, -1, svgData, null);
        }

        // Save the workbook with the inserted pictures and shapes
        workbook.Save("OutputWithPicturesAndShapes.xlsx");
    }
}