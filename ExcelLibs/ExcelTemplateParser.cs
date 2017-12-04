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
        private ExcelWorksheet _templateSheet;
        private ExcelWorksheet _outputSheet;

        public ExcelTemplateParser(ExcelWorksheet templateSheet,ExcelWorksheet outputSheet, Hash dataHash)
        {
            _templateSheet = templateSheet;
            _outputSheet = outputSheet;
            _dataHash = dataHash;
            _lastUsedRow = GetLastUsedRow();
            _lastUsedColumn = templateSheet.Dimension.End.Column;
        }

        public void Render()
        {
            var templateString = ParseToTemplateText();
            var dataString = AppliedDataToTemplateText(templateString);
            var elements = ParseToExcelElements(dataString);
            Apply(_templateSheet, _outputSheet, elements);
        }

        private IEnumerable<IExcelElement> ParseToExcelElements(string excelDataString)
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

        private String AppliedDataToTemplateText(string templateString)
        {
            Template template = Template.Parse(templateString);
            return template.Render(_dataHash);
        }

        private String ParseToTemplateText()
        {
            StringBuilder stringBuilder = new StringBuilder();
            Stack<String> forloopAddressStack = new Stack<string>();
            for (int r = 1; r <= _lastUsedRow; r++)
            {
                for (int c = 1; c <= _lastUsedColumn; c++)
                {
                    var cell = _templateSheet.Cells[r, c];

                    if (cell.Value == null) continue;

                    var cellValue = cell.Value.ToString().Replace("\n", "\\n");
                    if (Utils.IsForloopStart(cellValue))
                    {
                        ExcelAddress forLoopAddress = GetForLoopAddress(new ExcelCellAddress(r, c));
                        cellValue = cellValue.EndsWith(" %}") ? cellValue.Replace(" %}", " -%}") : cellValue;
                        stringBuilder.Append(cellValue + "\n");
                        stringBuilder.Append($"{ForloopElement.StartMarker}{forLoopAddress.Address}\n");
                        forloopAddressStack.Push(forLoopAddress.Address);
                    }
                    else if (Utils.IsForloopEnd(cellValue))
                    {
                        var lastForloopAddress = forloopAddressStack.Pop();
                        stringBuilder.Append($"{lastForloopAddress}{ForloopElement.EndMarker}\n");
                        cellValue = cellValue.EndsWith(" %}") ? cellValue.Replace(" %}", " -%}") : cellValue;
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

        private void Apply(ExcelWorksheet templateSheet, ExcelWorksheet outputSheet, IEnumerable<IExcelElement> elements)
        {
            int rowOffset = 0;
            foreach (var element in elements)
            {
                int rowOffsetInner = element.SetupSpace(templateSheet, outputSheet, rowOffset);
                element.ApplyStyle(templateSheet, outputSheet, rowOffset);
                element.ApplyValue(templateSheet, outputSheet, rowOffset);
                rowOffset += rowOffsetInner;
            }
        }

        private ExcelAddress GetForLoopAddress(ExcelCellAddress startAddress)
        {
            ExcelAddress forLoopEndAddress = null;
            int numOfInnerForLoop = 0;

            for (int r = startAddress.Row + 1; r <= _templateSheet.Dimension.End.Row; r++)
            {
                var cellValue = _templateSheet.GetValue<String>(r, startAddress.Column);
                if (cellValue != null)
                {
                    if (Utils.IsForloopStart(cellValue))
                    {
                        numOfInnerForLoop++;
                    }
                    else if (Utils.IsForloopEnd(cellValue))
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
            var row = _templateSheet.Dimension.End.Row;
            while (row >= 1)
            {
                var range = _templateSheet.Cells[row, 1, row, _templateSheet.Dimension.End.Column];
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
            foreach (string mergedCell in _templateSheet.MergedCells)
            {
                var mergeredAddress = new ExcelAddress(mergedCell);
                if (mergeredAddress.Start.Row == startAddress.Row &&
                    mergeredAddress.Start.Column == startAddress.Column)
                {
                    address = new ExcelAddress(mergedCell);
                    break;
                }
            }

            return address ?? new ExcelAddress(
                       startAddress.Row,
                       startAddress.Column,
                       startAddress.Row,
                       startAddress.Column);
        }
    }
}