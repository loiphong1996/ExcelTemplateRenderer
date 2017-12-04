using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DotLiquid;
using ExcelLibs.ExcelElement;
using OfficeOpenXml;

namespace ExcelLibs
{
    public class ExcelTemplateParser
    {
        private int _lastUsedRow;
        private int _lastUsedColumn;
        private Hash _dataHash;
        private ExcelWorksheet _sheet;

        public ExcelTemplateParser(ExcelWorksheet sheet, Hash dataHash)
        {
            _sheet = sheet;
            _dataHash = dataHash;
            _lastUsedRow = GetLastUsedRow();
            _lastUsedColumn = sheet.Dimension.End.Column;
        }

        public IEnumerable<IExcelElement> ParseToExcelElements(string excelDataString)
        {
            List<IExcelElement> result = new List<IExcelElement>();
            Queue<String> tokens = new Queue<String>(excelDataString.Split('\n'));
            while (tokens.Count != 0)
            {
                if (tokens.Peek().StartsWith(ForloopElement.StartMarker))
                {
                    result.Add(new ForloopCollection(tokens));
                }
                else
                {
                    result.Add(new ValueElement(tokens.Dequeue()));
                }
            }               

            return result;
        }

        public String AppliedDataToTemplateText(string templateString)
        {
            Template template = Template.Parse(templateString);
            return template.Render(_dataHash);
        }

        public String ParseToTemplateText()
        {
            StringBuilder stringBuilder = new StringBuilder();
            Stack<String> forloopAddressStack = new Stack<string>();
            for (int r = 1; r <= _lastUsedRow; r++)
            {
                for (int c = 1; c <= _lastUsedColumn; c++)
                {
                    var cell = _sheet.Cells[r, c];
                    
                    if(cell.Value == null) continue;
                    
                    var cellValue = cell.Value.ToString().Replace("\n","\\n");
//                    if (cell.Text.TrimStart().StartsWith("{% for") && cell.Text.TrimEnd().EndsWith("-%}"))
                    if (cell.Text.TrimStart().StartsWith("{% for"))
                    {
                        ExcelAddress forLoopAddress = GetForLoopAddress(new ExcelCellAddress(r, c));
                        stringBuilder.Append(cellValue + "\n");
                        stringBuilder.Append($"{ForloopElement.StartMarker}{forLoopAddress.Address}\n");
                        forloopAddressStack.Push(forLoopAddress.Address);
                    }
                    else if (cell.Text.Trim().StartsWith("{% endfor"))
                    {
                        var lastForloopAddress = forloopAddressStack.Pop();
                        stringBuilder.Append($"{lastForloopAddress}{ForloopElement.EndMarker}\n");
                        stringBuilder.Append(cellValue + "\n");
                    }
                    else if (!String.IsNullOrEmpty(cell.Text))
                    {
                        stringBuilder.Append($"{cell.Address},{cellValue}\n");
                    }
                }
            }

            return stringBuilder.ToString().TrimEnd();
        }

        public int Apply(ExcelWorksheet templateSheet,ExcelWorksheet outputSheet,IEnumerable<IExcelElement> elements)
        {
            int rowOffset = 0;
            foreach (var element in elements)
            {
                int rowOffsetInner = element.SetupSpace(templateSheet,outputSheet,rowOffset);
                element.ApplyStyle(templateSheet,outputSheet,rowOffset);
                element.ApplyValue(templateSheet,outputSheet,rowOffset);
                rowOffset += rowOffsetInner;
            }

            return rowOffset;
        }

        private ExcelAddress GetForLoopAddress(ExcelCellAddress startAddress)
        {
            ExcelAddress forLoopEndAddress = null;
            int numOfInnerForLoop = 0;

            for (int r = startAddress.Row + 1; r <= _sheet.Dimension.End.Row; r++)
            {
                var cellValue = _sheet.GetValue<String>(r, startAddress.Column);
                if (cellValue != null)
                {
                    if (cellValue.TrimStart().StartsWith("{% for"))
                    {
                        numOfInnerForLoop++;
                    }
                    else if (cellValue.TrimStart().StartsWith("{% endfor"))
                    {
                        if (numOfInnerForLoop == 0)
                        {
                            forLoopEndAddress =
                                GetMergeCellAddress(new ExcelCellAddress(r, startAddress.Column));
                            break;
                        }

                        numOfInnerForLoop--;
                    }
                }
            }

            if (forLoopEndAddress == null)
            {
                throw new InvalidOperationException("Can't not find complete for loop in with start address " +
                                                    startAddress);
            }

            return new ExcelAddress(
                startAddress.Row,
                startAddress.Column,
                forLoopEndAddress.End.Row,
                forLoopEndAddress.End.Column);
        }

        private int GetLastUsedRow()
        {
            var row = _sheet.Dimension.End.Row;
            while (row >= 1)
            {
                var range = _sheet.Cells[row, 1, row, _sheet.Dimension.End.Column];
                if (range.Any(cell => !string.IsNullOrEmpty(cell.Text)))
                {
                    break;
                }

                row--;
            }

            return row;
        }

        private ExcelAddress GetMergeCellAddress(ExcelCellAddress startAddress)
        {
            ExcelAddress address = null;
            foreach (string mergedCell in _sheet.MergedCells)
            {
                var mergeredAddress = new ExcelAddress(mergedCell);
                if (mergeredAddress.Start.Row == startAddress.Row &&
                    mergeredAddress.Start.Column == startAddress.Column)
                {
                    address = new ExcelAddress(mergedCell);
                    break;
                }
            }

            if (address == null)
            {
                address = new ExcelAddress(
                    startAddress.Row,
                    startAddress.Column,
                    startAddress.Row,
                    startAddress.Column);
            }

            return address;
        }
    }
}