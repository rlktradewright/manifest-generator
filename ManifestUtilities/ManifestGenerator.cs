using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace TradeWright.ManifestUtilities
{
    public class ManifestGenerator
    {
        private int mMajorVersion;
        private int mMinorVersion;
        private int mRevisionVersion;
        private string mExeName;
        private string mDescription;


        /// <summary>
        /// Generates an application manifest for the Visual Basic 6  project
        /// specified in <code>projectFilename</code>.
        /// </summary>
        /// 
        /// <param name="projectFilename">
        /// The path and filename of the required Visual Basic 6 project.
        /// </param>
        /// <returns></returns>
        public string Generate(String projectFilename, bool useVersion6CommonControls)
        {
            var sb = new StringBuilder();

            using (StreamReader sr = new StreamReader(projectFilename))
            {
                while (!sr.EndOfStream)
                {
                    processLine(sr.ReadLine(), sb, useVersion6CommonControls);
                }
            }

            return String.Format(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<assembly xmlns=""urn:schemas-microsoft-com:asm.v1"" manifestVersion=""1.0"" xmlns:asmv3=""urn:schemas-microsoft-com:asm.v3"">
    <assemblyIdentity name=""{0}"" processorArchitecture=""X86"" type=""win32"" version=""{1}"" />
    <description>{2}</description>
{3}</assembly>", mExeName, mMajorVersion.ToString() + "." + mMinorVersion.ToString() + ".0." + mRevisionVersion.ToString(), mDescription, sb.ToString());
        }


        private void processLine(string line, StringBuilder sb, bool useVersion6CommonControls)
        {
            if (String.IsNullOrEmpty(line)) return;

            var index = line.IndexOf('=');
            if (index < 0) return;

            var lineType = line.Substring(0, index);
            var lineContent = line.Substring(index + 1, line.Length - index - 1);

            Console.WriteLine(line);

            switch (lineType)
            {
            case "Description":
                mDescription = lineContent.Substring(1, lineContent.Length - 2);
                break;
            case "ExeName32":
                mExeName = lineContent.Substring(1, lineContent.Length - 2);
                break;
            case "MajorVer":
                mMajorVersion = Int32.Parse(lineContent);
                break;
            case "MinorVer":
                mMinorVersion = Int32.Parse(lineContent);
                break;
            case "Object":
                processObjectLine(lineContent, sb, useVersion6CommonControls);
                break;
            case "Reference":
                processReferenceLine(lineContent, sb);
                break;
            case "RevisionVer":
                mRevisionVersion = Int32.Parse(lineContent);
                break;
            case "Type":
                processTypeLine(lineContent, sb);
                break;
            }
        }

        private static void processObjectLine(string lineContent, StringBuilder sb, bool useVersion6CommonControls)
        {
            string guid = getGuid(lineContent);

            // if mscomctl, we need to generate a <dependentAssembly> element for the mscomctl.dll as well as
            // for the .ocx if version 6 is required
            if (guid.Equals("{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}", StringComparison.CurrentCultureIgnoreCase))
            {
                if (useVersion6CommonControls) outputDependentAssembly(sb, "Microsoft.Windows.Common-Controls", "6.0.0.0", "6595b64144ccf1df");
            }

            string version = getObjectTypelibVersion(lineContent);
            string objectFileName = getObjectFilename(guid, version);

            generateDependentAssembly(objectFileName, sb);
        }

        private static void processReferenceLine(string lineContent, StringBuilder sb)
        {
            string guid = getGuid(lineContent);

            // if stdole2, just ignore it since this is never side-by-sided
            if (guid.Equals("{00020430-0000-0000-C000-000000000046}", StringComparison.CurrentCultureIgnoreCase)) return;
            // if Microsoft Scripting Runtime, just ignore it since this is never side-by-sided
            if (guid.Equals("{420B2830-E718-11CF-893D-00A0C9054228}", StringComparison.CurrentCultureIgnoreCase)) return;

            string referenceFileName = getReferenceFilename(guid);

            if (referenceFileName.EndsWith(".tlb"))
            {
                outputTypelibFileInfo(sb, guid, referenceFileName);
            }
            else
            {
                generateDependentAssembly(referenceFileName, sb);
            }
        }

        private static void outputTypelibFileInfo(StringBuilder sb, string guid, string objectFileName)
        {
            sb.AppendLine(@"    <file name=""" + getFilenameFromPath(objectFileName) + @""">");
            sb.AppendLine(@"        <typelib tlbid=""" + guid + @""" version=""1.0"" flags=""hasdiskimage"" helpdir="""" />");
            sb.AppendLine(@"    </file>");
        }

        private static void generateDependentAssembly(string objectFileName, StringBuilder sb)
        {
            FileVersionInfo versionInfo = FileVersionInfo.GetVersionInfo(objectFileName);

            outputDependentAssembly(sb, getFilenameFromPath(versionInfo.FileName), versionInfo.FileVersion, "");
        }

        private static string getFilenameFromPath(string filepath)
        {
            string pattern = @"^([^\\]+\\)*(([^\.]+\.)+[^\.]+)$";
            var match = Regex.Match(filepath, pattern);
            if (!match.Success) throw new InvalidOperationException(String.Format("Filename not found in string {0}", filepath));
            return match.Groups[2].Value;
        }

        private static void outputDependentAssembly(StringBuilder sb, string fileName, string version, string publicKeyToken)
        {
            if (publicKeyToken != String.Empty) publicKeyToken = @"publicKeyToken=""" + publicKeyToken + @"""";
            sb.AppendLine("    <dependency>");
            sb.AppendLine("        <dependentAssembly>");
            sb.AppendLine(String.Format(@"            <assemblyIdentity name=""{0}"" processorArchitecture=""X86"" type=""win32"" version=""{1}"" {2}/>", fileName, version, publicKeyToken));
            sb.AppendLine(@"        </dependentAssembly>");
            sb.AppendLine(@"    </dependency>");
        }

        private static string getObjectTypelibVersion(string line)
        {
            string pattern = @"\#([0-9]+)\.([0-9]+)\#";
            var match = Regex.Match(line, pattern);

            if (!match.Success) throw new InvalidOperationException(String.Format("No typelib version found in string {1}", line));
            return numericStringToHex(match.Groups[1].Value) + "." + numericStringToHex(match.Groups[2].Value);
        }

        private static string numericStringToHex(string value)
        {
            return int.Parse(value).ToString("X");
        }

        private static string getObjectFilename(string guid, string version)
        {
            return (string)Registry.GetValue(@"HKEY_CLASSES_ROOT\Typelib\" + guid + @"\" + version + @"\0\win32", "", "");
        }

        private static string getReferenceFilename(string guid)
        {
            string fileName = String.Empty;
            var regKey = Registry.ClassesRoot.OpenSubKey(@"Typelib\" + guid);
            foreach (string subKeyName in regKey.GetSubKeyNames())
            {
                string s = (string)regKey.OpenSubKey(subKeyName + @"\0\win32").GetValue("");
                if (s != String.Empty) fileName = s;
            }

            if (fileName == String.Empty) throw new InvalidOperationException("Can't find filename for guid " + guid);
            return fileName;
        }

        private static string getGuid(string line)
        {
            // include the surrounding braces
            string pattern = @"(?i)\{([0-9]|[A-F]){8}\-([0-9]|[A-F]){4}\-([0-9]|[A-F]){4}\-([0-9]|[A-F]){4}\-([0-9]|[A-F]){12}\}";
            var match = Regex.Match(line, pattern);

            if (!match.Success) throw new InvalidOperationException(String.Format("No GUID found in string {0}", line));
            return match.Groups[0].Value;
        }

        private static void processTypeLine(string lineContent, StringBuilder sb)
        {
            if (!lineContent.Equals("Exe") &&
                !lineContent.Equals("Control") &&
                !lineContent.Equals("OleDll")) throw new ArgumentException("Wrong project type");
        }
    }
}
