using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace TradeWright.ManifestUtilities
{
    internal static class Utils
    {
        internal static string numericStringToHex(string value)
        {
            return int.Parse(value).ToString("X");
        }

        internal static string trimDelimiters(string value)
        {
            return value.Substring(1, value.Length - 2);
        }

    }
}
