using System;
using ExcelLibs.ExcelElement;
using OfficeOpenXml;

namespace ExcelLibs
{
    public class ValueElement : IExcelElement
    {
        private ExcelCellAddress _templateAddress;
        private String _value;

        public ExcelCellAddress TemplateAddress => _templateAddress;

        public ValueElement(ExcelCellAddress templateAddress, String value)
        {
            _templateAddress = templateAddress;
            _value = value;
        }

        public ValueElement(string token)
        {
            var separateIndex = token.IndexOf(',');
            if (separateIndex > 0)
            {
                _templateAddress = new ExcelCellAddress(token.Substring(0, separateIndex));
                _value = token.Substring(separateIndex + 1);
            }
        }

        public int SetupSpace(ExcelWorksheet templateSheet, ExcelWorksheet outputSheet, int rowOffset)
        {
            return 0;
        }

        public int ApplyValue(ExcelWorksheet templateSheet, ExcelWorksheet outputSheet, int rowOffset)
        {
            outputSheet.SetValue(
                _templateAddress.Row + rowOffset,
                _templateAddress.Column,
                _value.Replace("\\n", "\n"));
            return 0;
        }

        public void ApplyStyle(ExcelWorksheet templateSheet, ExcelWorksheet outputSheet, int rowOffset)
        {
            outputSheet.Cells[_templateAddress.Row + rowOffset, _templateAddress.Column].StyleID =
                templateSheet.Cells[_templateAddress.Row, _templateAddress.Column].StyleID;
            ExcelAddress templateMergedAddress = GetMergedAddress(templateSheet);
            if (templateMergedAddress != null)
            {
                ExcelAddress outputMergedAddress = new ExcelAddress(
                    templateMergedAddress.Start.Row + rowOffset,
                    templateMergedAddress.Start.Column,
                    templateMergedAddress.End.Row + rowOffset,
                    templateMergedAddress.End.Column);
                outputSheet.Cells[outputMergedAddress.Address].Merge = true;
            }

        }

        private ExcelAddress GetMergedAddress(ExcelWorksheet templateSheet)
        {
            foreach (string mergedCellAdress in templateSheet.MergedCells)
            {
                if (mergedCellAdress.Split(':')[0] == _templateAddress.Address)
                {
                    return new ExcelAddress(mergedCellAdress);
                }
            }

            return null;
        }
    }
}