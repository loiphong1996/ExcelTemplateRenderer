using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DotLiquid;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.Style;

namespace ExcelLibs
{
    public class ExcelRenderer
    {
        private int _currentForLoopRow = 0;
        private int _rowOffset = 0;
        private Dictionary<string, bool> _forLoopMarker = new Dictionary<string, bool>();
        private ExcelWorksheet _templateSheet;
        private ExcelWorksheet _outputSheet;
        private Hash _dataHash;
        private int _lastUsedColumn;
        private int _lastUsedRow;

        public ExcelRenderer(ExcelWorksheet template, ExcelWorksheet output, Hash data)
        {
            _templateSheet = template;
            _outputSheet = output;
            _dataHash = data;
            _lastUsedRow = GetLastUsedRow(_outputSheet);
            _lastUsedColumn = _outputSheet.Dimension.End.Column;
        }

        public void Render()
        {
            string templateString = ParseExcelValueToText(_templateSheet);
            string dataString = AppliedDataToText(templateString, _dataHash);
            string[] tokens = dataString.Split('\n');

            for (int i = 0; i < tokens.Length; i++)
            {
                if (String.IsNullOrEmpty(tokens[i])) continue;

                int separatorIndex = tokens[i].IndexOf(',');
                if (separatorIndex > 0)
                {
                    ExcelCellAddress cellAddress = new ExcelCellAddress(tokens[i].Substring(0, separatorIndex));
                    string content = tokens[i].Substring(separatorIndex + 1);
                    ExcelAddress forLoopAddress = GetForLoopAddress(cellAddress);
                    if (forLoopAddress == null)
                    {
                        _outputSheet.SetValue(cellAddress.Address, content);
                    }
                    else
                    {
//                        ExcelAddress offsetForLoopAddress = GetOffsetAddress(forLoopAddress);
//                        ExcelAddress offsetCellAddress = GetOffsetAddress(new ExcelAddress(cellAddress.Address));
//                        _outputSheet.SetValue(offsetForLoopAddress.Start.Row,cellAddress.Column,content);
                    }
                }
                else
                {
                    ExcelAddress templateForLoopAddress = new ExcelAddress(tokens[i]);
                    if (!_forLoopMarker[templateForLoopAddress.Address])
                    {
                        if (!IsInnerForloop(templateForLoopAddress))
                        {
                            _currentForLoopRow = templateForLoopAddress.Start.Row;                            
                        }
                    }
                    else
                    {
                        _currentForLoopRow +=
                            templateForLoopAddress.Rows - GetContentRows(templateForLoopAddress);
                    }

                    ExcelAddress offsetForLoopAddress = GetOffsetAddress(templateForLoopAddress);
                    SetupForloopChunk(templateForLoopAddress, offsetForLoopAddress);
                }
            }
        }

        private String AppliedDataToText(string templateString, Hash data)
        {
            Template template = Template.Parse(templateString);
            return template.Render(data);
        }

        private ExcelAddress GetForLoopAddress(ExcelCellAddress innerCell)
        {
            foreach (String addressString in _forLoopMarker.Keys)
            {
                ExcelAddress forLoopAddress = new ExcelAddress(addressString);
                if (forLoopAddress.Start.Column <= innerCell.Column &&
                    forLoopAddress.End.Column >= innerCell.Column &&
                    forLoopAddress.Start.Row < innerCell.Row &&
                    forLoopAddress.End.Row > innerCell.Row)
                {
                    return forLoopAddress;
                }
            }

            return null;
        }

        private ExcelAddress GetOffsetAddress(ExcelAddress address)
        {
            if (_forLoopMarker.ContainsKey(address.Address))
            {
                return new ExcelAddress(
                    _currentForLoopRow,
                    address.Start.Column,
                    address.End.Row + (address.Start.Row - _currentForLoopRow),
                    address.End.Column);
            }

            if (address.Start.Row < _currentForLoopRow)
            {
                return new ExcelAddress(
                    address.Start.Row + _rowOffset,
                    address.Start.Column,
                    address.End.Row + _rowOffset,
                    address.End.Column);
            }

            return address;
        }

        private void InsertRow(int row, int rows = 1)
        {
            _outputSheet.InsertRow(row, rows);
            _rowOffset = _rowOffset + rows;
        }

        private void RemoveRow(int row, int rows = 1)
        {
            _outputSheet.DeleteRow(row, rows, true);
            _rowOffset = _rowOffset - rows;
        }

        private void SetupForloopChunk(ExcelAddress templateChunk, ExcelAddress outputChunk)
        {
            int contentRows = GetContentRows(templateChunk);
            if (!_forLoopMarker[templateChunk.Address]) //check if already remove template chunk of output file
            {
                RemoveRow(outputChunk.Start.Row, contentRows + 2);
                _forLoopMarker[templateChunk.Address] = true;
            }

            InsertRow(outputChunk.Start.Row, contentRows);
//            _templateSheet.Cells[
//                    templateChunk.Start.Row + 1,
//                    templateChunk.Start.Column,
//                    templateChunk.End.Row - 1,
//                    templateChunk.End.Column]
//                .Copy(_outputSheet.Cells[
//                    outputChunk.Start.Row + 1,
//                    outputChunk.Start.Column,
//                    outputChunk.End.Row - 1,
//                    outputChunk.End.Column]);
        }

        private int GetContentRows(ExcelAddress templateForloopChunk)
        {
            int contentRows = 0;
            bool isInInnerForLoopChunk = false;
            foreach (var cell in _templateSheet.Cells[
                templateForloopChunk.Start.Row + 1,
                templateForloopChunk.Start.Column,
                templateForloopChunk.End.Row - 1,
                templateForloopChunk.Start.Column])
            {
                if (cell.Text.TrimStart().StartsWith("{% for") && cell.Text.TrimEnd().EndsWith("-%}"))
                {
                    isInInnerForLoopChunk = true;
                    continue;
                }

                if (cell.Text.Trim().StartsWith("{% endfor -%}"))
                {
                    isInInnerForLoopChunk = false;
                    continue;
                }

                if (!isInInnerForLoopChunk)
                {
                    contentRows++;
                }
            }

            return contentRows;
        }

        private bool IsInnerForloop(ExcelAddress forLoopChunk)
        {
            foreach (string addressString in _forLoopMarker.Keys)
            {
                var address = new ExcelAddress(addressString);
                if (address.Start.Row < forLoopChunk.Start.Row &&
                    address.Start.Column <= forLoopChunk.Start.Column &&
                    address.End.Row > forLoopChunk.End.Row &&
                    address.End.Column >= forLoopChunk.End.Column)
                {
                    return true;
                }
            }

            return false;
        }

        private String ParseExcelValueToText(ExcelWorksheet sheet)
        {
            StringBuilder stringBuilder = new StringBuilder();
            for (int r = 1; r <= _lastUsedRow; r++)
            {
                for (int c = 1; c <= _lastUsedColumn; c++)
                {
                    var cell = sheet.Cells[r, c];
                    if (cell.Text.TrimStart().StartsWith("{% for") && cell.Text.TrimEnd().EndsWith("-%}"))
                    {
                        ExcelAddress forLoopAddress = GetForLoopAddress(sheet, new ExcelCellAddress(r, c));
                        _forLoopMarker.Add(forLoopAddress.Address, false);
                        stringBuilder.Append(cell.Text + "\n");
                        stringBuilder.Append($"{forLoopAddress.Address}\n");
                    }
                    else if (cell.Text.Trim().StartsWith("{% endfor -%}"))
                    {
                        stringBuilder.Append(cell.Text + "\n");
                    }
                    else if (!String.IsNullOrEmpty(cell.Text))
                    {
                        stringBuilder.Append($"{cell.Address},{cell.Text}\n");
                    }
                }
            }

            return stringBuilder.ToString();
        }

        private ExcelAddress GetForLoopAddress(ExcelWorksheet sheet, ExcelCellAddress startAddress)
        {
            ExcelAddress forLoopEndAddress = null;
            int numOfInnerForLoop = 0;

            for (int r = startAddress.Row + 1; r <= sheet.Dimension.End.Row; r++)
            {
                var cellValue = sheet.GetValue<String>(r, startAddress.Column);
                if (cellValue != null)
                {
                    if (cellValue.TrimStart().StartsWith("{% for") && cellValue.TrimEnd().EndsWith("-%}"))
                    {
                        numOfInnerForLoop++;
                    }
                    else if (cellValue.Trim().Equals("{% endfor -%}"))
                    {
                        if (numOfInnerForLoop == 0)
                        {
                            forLoopEndAddress =
                                GetMergeCellAddress(sheet, new ExcelCellAddress(r, startAddress.Column));
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

        private ExcelAddress GetMergeCellAddress(ExcelWorksheet sheet, ExcelCellAddress startAddress)
        {
            ExcelAddress address = null;
            foreach (string mergedCell in sheet.MergedCells)
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

        private int GetLastUsedRow(ExcelWorksheet sheet)
        {
            var row = sheet.Dimension.End.Row;
            while (row >= 1)
            {
                var range = sheet.Cells[row, 1, row, sheet.Dimension.End.Column];
                if (range.Any(cell => !string.IsNullOrEmpty(cell.Text)))
                {
                    break;
                }

                row--;
            }

            return row;
        }
    }
}