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
                _value.Replace("\\n","\n"));
            return 0;
        }

        public void ApplyStyle(ExcelWorksheet templateSheet, ExcelWorksheet outputSheet, int rowOffset)
        {
            CopyStyle(templateSheet.Cells[_templateAddress.Row, _templateAddress.Column],
                outputSheet.Cells[_templateAddress.Row + rowOffset, _templateAddress.Column]);
//            templateSheet.Cells[_templateAddress.Row,_templateAddress.Column].Copy(
//                );
        }

        private void CopyStyle(ExcelRange templateCell, ExcelRange outputCell)
        {
            outputCell.StyleID = templateCell.StyleID;
        }
    }
}