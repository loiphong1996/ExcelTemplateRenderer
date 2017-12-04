using System;
using System.Collections.Generic;
using ExcelLibs.ExcelElement;
using OfficeOpenXml;

namespace ExcelLibs
{
    public class ForloopElement : IExcelElement
    {
        public static readonly string StartMarker = "<<--";
        public static readonly string EndMarker = "-->>";
        private ExcelAddress _templateAddress;
        private LinkedList<IExcelElement> _elements = new LinkedList<IExcelElement>();
        private int _contentRows = -1;

        public ExcelAddress TemplateAddress => _templateAddress;


        public ForloopElement(String templateAddress)
        {
            _templateAddress = new ExcelAddress(templateAddress);
        }

        public ForloopElement(Queue<String> tokens)
        {
            while (tokens.Count != 0)
            {
                if (tokens.Peek().StartsWith(StartMarker))
                {
                    if (_templateAddress == null)
                    {
                        _templateAddress = new ExcelAddress(tokens.Dequeue().Substring(StartMarker.Length));
                    }
                    else
                    {
                        _elements.AddLast(new ForloopCollection(tokens));
                    }
                }
                else if (tokens.Peek().EndsWith(EndMarker))
                {
                    tokens.Dequeue();
                    break;
                }
                else
                {
                    _elements.AddLast(new ValueElement(tokens.Dequeue()));
                }
            }
        }

        public void Add(IExcelElement element)
        {
            _elements.AddLast(element);
        }

        public bool IsCanContain(ValueElement element)
        {
            return _templateAddress.Start.Row < element.TemplateAddress.Row &&
                   _templateAddress.Start.Column <= element.TemplateAddress.Column &&
                   _templateAddress.End.Row > element.TemplateAddress.Row &&
                   _templateAddress.End.Column >= element.TemplateAddress.Column;
        }

        public int CreateSpace(ExcelWorksheet templateSheet, ExcelWorksheet outputSheet)
        {
            var contentRows = GetContentRows(templateSheet);
            outputSheet.InsertRow(_templateAddress.Start.Row, contentRows);
            return contentRows;
        }

        //TODO case which same row have multilple for loop
        public int GetContentRows(ExcelWorksheet templateSheet)
        {
            if (_contentRows >= 0)
            {
                return _contentRows;
            }

            bool innerForLoopMarker = false;
            int contentRows = 0;

            for (int r = _templateAddress.Start.Row + 1; r <= _templateAddress.End.Row - 1; r++)
            {
                for (int c = _templateAddress.Start.Row; c <= _templateAddress.End.Column; c++)
                {
                    var cellValue = templateSheet.GetValue<String>(r, c);
                    if (cellValue != null)
                    {
                        if (Utils.IsForloopStart(cellValue))
                        {
                            innerForLoopMarker = true;
                            break;
                        }

                        if (Utils.IsForloopEnd(cellValue))
                        {
                            innerForLoopMarker = false;
                            break;
                        }
                    }
                }

                if (innerForLoopMarker == false)
                {
                    contentRows++;
                }
            }

            _contentRows = contentRows;
            return contentRows;
        }

        public int SetupSpace(ExcelWorksheet templateSheet, ExcelWorksheet outputSheet, int rowOffset)
        {
            return CreateSpace(templateSheet, outputSheet);
        }

        public int ApplyValue(ExcelWorksheet templateSheet, ExcelWorksheet outputSheet, int rowOffset)
        {
            foreach (IExcelElement excelElement in _elements)
            {
                excelElement.ApplyValue(templateSheet, outputSheet, rowOffset);
            }

            return 0;
        }

        public void ApplyStyle(ExcelWorksheet templateSheet, ExcelWorksheet outputSheet, int rowOffset)
        {
            for (int r = _templateAddress.Start.Row + 1; r <= _templateAddress.End.Row - 1; r++)
            {
                for (int c = _templateAddress.Start.Column; c <= _templateAddress.End.Column; c++)
                {
                    ExcelRangeBase templateRange = templateSheet.Cells[r, c];
                    ExcelRange outputRange = outputSheet.Cells[
                        templateRange.Start.Row + rowOffset,
                        templateRange.Start.Column,
                        templateRange.End.Row + rowOffset,
                        templateRange.End.Column
                    ];
                    outputRange.StyleID = templateRange.StyleID;

                    ExcelAddress templateMergedAddress = GetMergedAddress(templateSheet, new ExcelCellAddress(r, c));
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
            }
        }

        private ExcelAddress GetMergedAddress(ExcelWorksheet templateSheet, ExcelCellAddress cellAddress)
        {
            foreach (string mergedCellAdress in templateSheet.MergedCells)
            {
                if (mergedCellAdress.Split(':')[0] == cellAddress.Address)
                {
                    return new ExcelAddress(mergedCellAdress);
                }
            }

            return null;
        }
    }
}