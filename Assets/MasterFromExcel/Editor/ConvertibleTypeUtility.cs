using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace MasterFromExcel
{
    public static class ConvertibleTypeUtility
    {
        static readonly IEnumerable<string> ConvertibleTypes = new[]
        {
            "int", "long", "float", "double", "bool", "string", "datetime", "enum",
        };

        public static bool IsDefined(string type)
        {
            return ConvertibleTypes.Contains(Regex.Replace(type.ToLower(), @"\[\]$", ""));
        }

        public static string GetString(string type)
        {
            return IsDefined(type) ? type.ToLower() : $"{type.ToTopUpper()}Key";
        }
    }

}
