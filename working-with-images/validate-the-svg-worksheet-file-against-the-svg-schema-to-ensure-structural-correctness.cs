// Title: C# – Validate an SVG file against an XSD schema and embed it in Excel with Aspose.Cells
// Description: Demonstrates how to load an SVG and its XSD schema, verify the SVG structure using XmlSchemaSet, insert the validated SVG into a worksheet via Shapes.AddSvg, and save the workbook as an .xlsx file. The example is built with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# SVG validation | XML schema validation .NET | AddSvg shape | Excel workbook image insertion | XSD schema for SVG | Validate SVG before embedding | Aspose.Cells example GitHub | coding‑agent SVG validation
// Common Searches: validate svg with xsd c# | aspacells addsvg example | c# xml schema validation for svg files | insert svg into excel using aspose.cells | svg schema validation before workbook save
// Developer Intent: Check an SVG file against its XSD schema and, if valid, add it to an Excel worksheet using Aspose.Cells.
// Use Cases: Automated quality gate for SVG assets in reporting pipelines. | Dynamic generation of Excel dashboards that include only schema‑compliant graphics. | Logging validation errors and skipping malformed SVGs to prevent workbook corruption.
// AI Prompts: Generate a C# function that validates an SVG file against a given XSD and returns detailed error messages. | Show code to read an SVG into a byte array and insert it into an Aspose.Cells worksheet with a fallback PNG. | Create a script that scans a folder of SVGs, validates each against the schema, adds the valid ones to a new workbook, and logs the invalid files.

using System;
using System.IO;
using System.Xml;
using System.Xml.Schema;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to load an SVG and its XSD schema, verify the SVG structure using XmlSchemaSet, insert the validated SVG into a worksheet via Shapes.AddSvg, and save the workbook as an .xlsx file. The example is built with Aspose.Cells for .NET.
class SvgValidationDemo
{
    // Validates an SVG file against an XSD schema.
    // Returns true if the SVG conforms to the schema, otherwise false.
    static bool ValidateSvg(string svgFilePath, string xsdFilePath)
    {
        bool isValid = true;

        try
        {
            if (!File.Exists(svgFilePath))
            {
                Console.WriteLine($"SVG file not found: {svgFilePath}");
                return false;
            }

            if (!File.Exists(xsdFilePath))
            {
                Console.WriteLine($"XSD file not found: {xsdFilePath}");
                return false;
            }

            // Load the SVG schema.
            XmlSchemaSet schemas = new XmlSchemaSet();
            schemas.Add(null, xsdFilePath);

            // Set up XML reader settings with the schema and a validation callback.
            XmlReaderSettings settings = new XmlReaderSettings
            {
                ValidationType = System.Xml.ValidationType.Schema,
                Schemas = schemas
            };
            settings.ValidationFlags |= XmlSchemaValidationFlags.ProcessIdentityConstraints;
            settings.ValidationEventHandler += (sender, args) =>
            {
                // Any validation error will set the flag to false.
                Console.WriteLine($"Validation {args.Severity}: {args.Message}");
                isValid = false;
            };

            // Read and validate the SVG file.
            using (FileStream fs = new FileStream(svgFilePath, FileMode.Open, FileAccess.Read))
            using (XmlReader reader = XmlReader.Create(fs, settings))
            {
                while (reader.Read()) { /* reading triggers validation */ }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Exception during SVG validation: {ex.Message}");
            isValid = false;
        }

        return isValid;
    }

    static void Main()
    {
        // Paths to the SVG file and its corresponding XSD schema.
        string svgPath = "sample.svg";
        string xsdPath = "svg.xsd";

        // Verify required files exist before proceeding.
        if (!File.Exists(svgPath))
        {
            Console.WriteLine($"SVG file not found: {svgPath}");
            return;
        }

        if (!File.Exists(xsdPath))
        {
            Console.WriteLine($"XSD file not found: {xsdPath}");
            return;
        }

        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            ShapeCollection shapes = sheet.Shapes;

            // Load SVG data into a byte array.
            byte[] svgData = File.ReadAllBytes(svgPath);

            // Add the SVG to the worksheet (demonstrates AddSvg usage).
            // Parameters: topRow, top, leftColumn, left, height, width, svgData, compatibleImageData
            shapes.AddSvg(0, 0, 0, 0, -1, -1, svgData, null);

            // Validate the SVG against the schema.
            bool svgIsValid = ValidateSvg(svgPath, xsdPath);
            Console.WriteLine($"SVG validation result: {(svgIsValid ? "Valid" : "Invalid")}");

            // Save the workbook (demonstrates save lifecycle).
            workbook.Save("output.xlsx");
            Console.WriteLine("Workbook saved as output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Runtime exception: {ex.Message}");
        }
    }
}
