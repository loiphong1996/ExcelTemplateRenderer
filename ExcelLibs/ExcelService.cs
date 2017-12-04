using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Policy;
using System.Xml;
using OfficeOpenXml;
using DotLiquid;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using Hash = DotLiquid.Hash;


namespace ExcelLibs
{
    public class ExcelService
    {
        public void Render(FileInfo templateFile, FileInfo outputFile, Hash data)
        {
            using (var template = new ExcelPackage(templateFile))
            {
                using (var output = new ExcelPackage(outputFile, templateFile))
                {
                    for (int i = 1; i <= template.Workbook.Worksheets.Count; i++)
                    {
                        var templateSheet = template.Workbook.Worksheets[i];
                        var outputSheet = output.Workbook.Worksheets[i];
                        var parser = new ExcelTemplateParser(templateSheet,outputSheet, data);
                        parser.Render();
                    }
                    output.Save();
                }
            }
        }
    }
}