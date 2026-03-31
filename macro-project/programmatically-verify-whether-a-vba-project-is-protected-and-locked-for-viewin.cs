using System;
using System.Xml;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaVerification
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file (macro-enabled .xlsm)
            string workbookPath = "sample.xlsm";

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Retrieve protection status
            bool isProtected = vbaProject.IsProtected;
            bool isLockedForViewing = vbaProject.IslockedForViewing;

            // Create an XML document to hold the verification result
            XmlDocument xmlDoc = new XmlDocument();

            // Root element
            XmlElement root = xmlDoc.CreateElement("VbaProjectStatus");
            xmlDoc.AppendChild(root);

            // IsProtected element
            XmlElement protectedElem = xmlDoc.CreateElement("IsProtected");
            protectedElem.InnerText = isProtected.ToString();
            root.AppendChild(protectedElem);

            // IsLockedForViewing element
            XmlElement lockedElem = xmlDoc.CreateElement("IsLockedForViewing");
            lockedElem.InnerText = isLockedForViewing.ToString();
            root.AppendChild(lockedElem);

            // Save the XML to a file
            string xmlOutputPath = "VbaProjectStatus.xml";
            xmlDoc.Save(xmlOutputPath);

            // Optional: display result on console
            Console.WriteLine($"VBA Project Protected: {isProtected}");
            Console.WriteLine($"VBA Project Locked for Viewing: {isLockedForViewing}");
            Console.WriteLine($"Verification XML saved to: {xmlOutputPath}");
        }
    }
}