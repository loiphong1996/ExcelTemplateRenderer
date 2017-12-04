using System;
using System.Collections.Generic;
using System.Linq;
using ExcelLibs.ExcelElement;
using OfficeOpenXml;

namespace ExcelLibs
{
    public class ForloopCollection : IExcelElement
    {
        private LinkedList<ForloopElement> _forloopElements = new LinkedList<ForloopElement>();
        private int _forloopOffset = 0;

        public ForloopCollection(Queue<string> tokens)
        {
            while (tokens.Count != 0)
            {
                if (tokens.Peek().StartsWith(ForloopElement.StartMarker) &&
                    IsValidTemplateAddress(tokens.Peek().Substring(ForloopElement.StartMarker.Length)))
                {
                    _forloopElements.AddLast(new ForloopElement(tokens));
                }
                else
                {
                    break;
                }
            }
        }

        public void Add(ForloopElement element)
        {
            if (!IsValidTemplateAddress(element))
            {
                throw new InvalidOperationException(
                    "ForloopElements in one collection must have same template address!");
            }

            _forloopElements.AddLast(element);
        }

        public bool IsValidTemplateAddress(String addressString)
        {
            if (_forloopElements.Count > 0)
            {
                return _forloopElements.All(forloopElement =>
                    addressString == forloopElement.TemplateAddress.Address);
            }

            return true;
        }

        public bool IsValidTemplateAddress(ForloopElement element)
        {
            return _forloopElements.All(forloopElement =>
                element.TemplateAddress.Address == forloopElement.TemplateAddress.Address);
        }

        private int CleanupTemplateSpace(ExcelWorksheet templateSheet, ExcelWorksheet outputSheet)
        {
            ForloopElement templateElement = _forloopElements.First.Value;
            int rowsToBeDelete = templateElement.GetContentRows(templateSheet) + 2;

            outputSheet.DeleteRow(templateElement.TemplateAddress.Start.Row, rowsToBeDelete, true);
            _forloopOffset = -1;
            return -rowsToBeDelete;
        }

        public int SetupSpace(ExcelWorksheet templateSheet, ExcelWorksheet outputSheet, int rowOffset)
        {
            int offset = CleanupTemplateSpace(templateSheet, outputSheet);
            foreach (ForloopElement element in _forloopElements)
            {
                offset += element.SetupSpace(templateSheet, outputSheet, rowOffset);
            }

            return offset;
        }

        public int ApplyValue(ExcelWorksheet templateSheet, ExcelWorksheet outputSheet, int rowOffset)
        {            
            
            foreach (ForloopElement element in _forloopElements)
            {
                element.ApplyValue(templateSheet, outputSheet, rowOffset + _forloopOffset);
                element.ApplyStyle(templateSheet, outputSheet, rowOffset + _forloopOffset);
                _forloopOffset += element.GetContentRows(templateSheet);
            }
            
            return 0;
        }

        private void ApplyOuterStyle()
        {
            
        }
        

        public void ApplyStyle(ExcelWorksheet templateSheet, ExcelWorksheet outputSheet, int rowOffset)
        {
            //do nothing
        }
        
        
    }
}