using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

using TLI;

namespace TradeWright.ManifestUtilities
{
    public class ManifestGenerator
    {
        private string mType;
        private int mMajorVersion;
        private int mMinorVersion;
        private int mRevisionVersion;
        private string mObjectFileName;
        private string mObjectFilePath;
        private string mDescription;
        private string mProjectPath;

        [FlagsAttribute]
        private enum LIBFLAGS : short
        {
            Restricted = 1,
            Control = 2,
            Hidden = 4,
            HasDiskImage = 8
        }

        // note the capitalisation in these names is important
        // since they become XML element names
        private enum MISCSTATUS
        {
            miscStatus = 0,
            miscStatusContent = 1,
            miscStatusThumbnail = 2,
            miscStatusIcon = 3,
            miscStatusDocPrint = 4
        }

        [FlagsAttribute]
        private enum OLEMISC
        {
            RECOMPOSEONRESIZE = 0x1,
            ONLYICONIC = 0x2,
            INSERTNOTREPLACE = 0x4,
            STATIC = 0x8,
            CANTLINKINSIDE = 0x10,
            CANLINKBYOLE1 = 0x20,
            ISLINKOBJECT = 0x40,
            INSIDEOUT = 0x80,
            ACTIVATEWHENVISIBLE = 0x100,
            RENDERINGISDEVICEINDEPENDENT = 0x200,
            INVISIBLEATRUNTIME = 0x400,
            ALWAYSRUN = 0x800,
            ACTSLIKEBUTTON = 0x1000,
            ACTSLIKELABEL = 0x2000,
            NOUIACTIVATE = 0x4000,
            ALIGNABLE = 0x8000,
            SIMPLEFRAME = 0x10000,
            SETCLIENTSITEFIRST = 0x20000,
            IMEMODE = 0x40000,
            IGNOREACTIVATEWHENVISIBLE = 0x80000,
            WANTSTOMENUMERGE = 0x100000,
            SUPPORTSMULTILEVELUNDO = 0x200000
        }



        /// <summary>
        /// Generates a manifest for the Visual Basic 6  project
        /// specified in <code>projectFilename</code>.
        /// </summary>
        /// 
        /// <param name="projectFilename">
        /// The path and filename of the required Visual Basic 6 project.
        /// </param>
        /// <param name="useVersion6CommonControls"
        /// Indicates whether version 6 of the Windows Common Controls are 
        /// to be used. 
        /// 
        /// If the program does not use the Windows Common 
        /// Controls then this parameter is ignored.
        /// <returns></returns>
        public string Generate(String projectFilename, bool useVersion6CommonControls)
        {
            mProjectPath = Utils.getPathFromFilePath(projectFilename);
            mObjectFilePath = mProjectPath;

            var sb = new StringBuilder();

            processProjectFile(projectFilename, useVersion6CommonControls, sb);

            if (!mType.Equals("Exe"))
            {
                generateTypesInfo(mObjectFilePath + @"\" + mObjectFileName, sb);
            }

            return String.Format(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<assembly xmlns=""urn:schemas-microsoft-com:asm.v1"" manifestVersion=""1.0"" xmlns:asmv3=""urn:schemas-microsoft-com:asm.v3"">
    <assemblyIdentity name=""{0}"" processorArchitecture=""X86"" type=""win32"" version=""{1}"" />
    <description>{2}</description>
{3}</assembly>", 
                                mObjectFileName, 
                                mMajorVersion.ToString() + "." + mMinorVersion.ToString() + ".0." + mRevisionVersion.ToString(), 
                                mDescription, 
                                sb.ToString());
        }

        private void processProjectFile(String projectFilename, bool useVersion6CommonControls, StringBuilder sb)
        {
            using (StreamReader sr = new StreamReader(projectFilename))
            {
                while (!sr.EndOfStream)
                {
                    processLine(sr.ReadLine(), sb, useVersion6CommonControls);
                }
            }
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
                mDescription = Utils.trimDelimiters(lineContent);
                break;
            case "ExeName32":
                mObjectFileName = Utils.trimDelimiters(lineContent);
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
            case "Path32":
                mObjectFilePath = mProjectPath + Utils.trimDelimiters(lineContent);
                break;
            case "Reference":
                processReferenceLine(lineContent, sb);
                break;
            case "RevisionVer":
                mRevisionVersion = Int32.Parse(lineContent);
                break;
            case "Type":
                mType = processTypeLine(lineContent);
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

        private static string processTypeLine(string lineContent)
        {
            if (!lineContent.Equals("Exe") &&
                !lineContent.Equals("Control") &&
                !lineContent.Equals("OleDll")) throw new ArgumentException("Wrong project type");
            return lineContent;
        }

        private void generateTypesInfo(string assemblyFilename, StringBuilder sb)
        {
            sb.AppendLine(string.Format(@"    <file name=""{0}"">", Utils.getFilenameFromFilePath(assemblyFilename)));
            var tlia = new TLI.TLIApplication();
            var typelibInfo = tlia.TypeLibInfoFromFile(assemblyFilename);

            generateTypelibElement(typelibInfo, sb);
            generateComClassElements(typelibInfo, sb);
            generateComInterfaceExternalProxyStubElements(typelibInfo, sb);
            sb.AppendLine(@"    </file>");
        }

        private void generateTypelibElement(TLI.TypeLibInfo typelibInfo, StringBuilder sb)
        {
            sb.AppendLine(string.Format(@"        <typelib tlbid=""{0}"" version=""{1}"" flags=""{2}"" helpdir="""" />",
                                        typelibInfo.GUID,
                                        typelibInfo.MajorVersion.ToString() + "." + typelibInfo.MinorVersion.ToString(),
                                        Enum.Format(typeof(LIBFLAGS), (LIBFLAGS)typelibInfo.AttributeMask, "G")));
        }

        private void generateComClassElements(TypeLibInfo typelibInfo, StringBuilder sb)
        {
            foreach (CoClassInfo c in typelibInfo.CoClasses)
            {
                if (!String.IsNullOrEmpty((string)Registry.GetValue(@"HKEY_CLASSES_ROOT\ClsId\" + c.GUID + @"\InprocServer32", "ThreadingModel", "")))
                {
                    generateComClassElement(typelibInfo, sb, c);
                }
            }
        }

        private static void generateComClassElement(TypeLibInfo typelibInfo, StringBuilder sb, CoClassInfo c)
        {
            string progID = (string)Registry.GetValue(@"HKEY_CLASSES_ROOT\ClsId\" + c.GUID + @"\VersionIndependentProgID", "", "");
            if (String.IsNullOrEmpty(progID)) progID = (string)Registry.GetValue(@"HKEY_CLASSES_ROOT\ClsId\" + c.GUID + @"\ProgID", "", "" /*typelibInfo.Name + "." + c.Name*/);
            
            string curVer = (string)Registry.GetValue(@"HKEY_CLASSES_ROOT\" + progID + @"\CurVer", "", "");
            string progIDAtt = (!String.IsNullOrEmpty(curVer)) ? @" progid=""" + curVer + @"""" : String.Empty;

            string threadingModel = (string)Registry.GetValue(@"HKEY_CLASSES_ROOT\ClsId\" + c.GUID + @"\InprocServer32", "ThreadingModel", "");
            string threadingModelAtt = (!String.IsNullOrEmpty(threadingModel)) ? @" threadingModel=""" + threadingModel + @"""" : String.Empty;

            string miscStatus = String.Empty;
            foreach (int i in Enum.GetValues(typeof(MISCSTATUS)))
            {
                int flags;
                Int32.TryParse((string)Registry.GetValue(@"HKEY_CLASSES_ROOT\ClsId\" + c.GUID + @"\MiscStatus" + (string)(i > 0 ? @"\" + i.ToString() : String.Empty), "", ""), out flags);
                if (flags != 0)
                {
                    miscStatus = miscStatus + " " + Enum.GetName(typeof(MISCSTATUS), i) + @"=""" + Enum.Format(typeof(OLEMISC), flags, "G") + @"""";
                }
            }

            sb.Append(String.Format(@"        <comClass clsid=""{0}"" tlbid=""{1}""{2}{3}{4}", c.GUID, typelibInfo.GUID, progIDAtt, threadingModelAtt, miscStatus));

            if (progID.Equals(curVer))
            {
                sb.AppendLine(@"/>");
            }
            else
            {
                sb.AppendLine("\n" + @"            <progid>" + progID + @"</progid>");
                sb.AppendLine(@"        </comClass>");
            }
        }

        private void generateComInterfaceExternalProxyStubElements(TypeLibInfo typelibInfo, StringBuilder sb)
        {
            var interfaces = new Dictionary<string, InterfaceInfo>();
            foreach (CoClassInfo c in typelibInfo.CoClasses)
            {
                if (c.DefaultInterface != null)
                {
                    if (!interfaces.ContainsKey(c.DefaultInterface.GUID)) interfaces.Add(c.DefaultInterface.GUID, c.DefaultInterface);
                }
            }

            foreach (InterfaceInfo i in interfaces.Values)
            {
                generateComInterfaceExternalProxyStubElement(i, typelibInfo, sb);
            }
        }

        private void generateComInterfaceExternalProxyStubElement(InterfaceInfo i, TypeLibInfo typelibInfo, StringBuilder sb)
        {
            sb.AppendLine(String.Format(@"        <comInterfaceExternalProxyStub name=""{0}"" iid=""{1}"" tlbid=""{2}"" proxyStubClsid32=""{3}"" />",
                                        i.Name,
                                        i.GUID,
                                        typelibInfo.GUID,
                                        (string)Registry.GetValue(@"HKEY_CLASSES_ROOT\Interface\" + i.GUID + @"\ProxyStubClsid32", "", "")));
        }

        private static void outputTypelibFileInfo(StringBuilder sb, string guid, string objectFileName)
        {
            sb.AppendLine(@"    <file name=""" + Utils.getFilenameFromFilePath(objectFileName) + @""">");
            sb.AppendLine(@"        <typelib tlbid=""" + guid + @""" version=""1.0"" flags=""hasdiskimage"" helpdir="""" />");
            sb.AppendLine(@"    </file>");
        }

        private static void generateDependentAssembly(string objectFileName, StringBuilder sb)
        {
            outputDependentAssembly(sb, Utils.getFilenameFromFilePath(objectFileName), FileVersionInfo.GetVersionInfo(objectFileName).FileVersion, "");
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

            if (!match.Success) throw new InvalidOperationException(String.Format("No typelib version found in string {0}", line));
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
            using (var regKey = Registry.ClassesRoot.OpenSubKey(@"Typelib\" + guid))
            {
                foreach (string subKeyName in regKey.GetSubKeyNames())
                {
                    using (var subKey = regKey.OpenSubKey(subKeyName + @"\0\win32"))
                    {
                        string s = (string)subKey.GetValue("");
                        if (s != String.Empty) fileName = s;
                    }
                }
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

    }
}
