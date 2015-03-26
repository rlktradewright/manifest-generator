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
        internal static string getFilenameFromFilePath(string filepath)
        {
            string pattern = @"^([^\\]+\\)*(([^\.]+\.)+[^\.]+)$";
            var match = Regex.Match(filepath, pattern);
            if (!match.Success) throw new InvalidOperationException(String.Format("Filename not found in string {0}", filepath));
            return match.Groups[2].Value;
        }

        internal static string getPathFromFilePath(string filepath)
        {
            string pattern = @"^(([^\\]+\\)*)(([^\.]+\.)+[^\.]+)$";
            var match = Regex.Match(filepath, pattern);
            if (!match.Success) throw new InvalidOperationException(String.Format("Path not found in string {0}", filepath));
            return match.Groups[1].Value;
        }

        internal static string trimDelimiters(string value)
        {
            return value.Substring(1, value.Length - 2);
        }

    }
}
