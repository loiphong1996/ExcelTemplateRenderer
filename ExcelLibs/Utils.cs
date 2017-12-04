using System;

namespace ExcelLibs
{
    public class Utils
    {
        
        public static bool IsForloopStart(String value)
        {
            value = value.Trim();
            return value.StartsWith("{% for") &&
                   (value.EndsWith("-%}") || value.EndsWith("%}"));
        }
        public static bool IsForloopEnd(String value)
        {
            value = value.Trim();
            return value.StartsWith("{% endfor") &&
                   (value.EndsWith("-%}") || value.EndsWith("%}"));
        }
        
    }
}