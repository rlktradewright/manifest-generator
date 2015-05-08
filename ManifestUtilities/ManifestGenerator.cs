using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;

using TLI;

namespace TradeWright.ManifestUtilities
{
    public class ManifestGenerator
    {
        
        [FlagsAttribute]
        private enum LIBFLAGS : short
        {
            restricted = 1,
            control = 2,
            hidden = 4,
            hasDiskImage = 8
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
            recomposeOnResize = 0x1,
            onlyIconic = 0x2,
            insertNotReplace = 0x4,
            Static = 0x8,
            cantLinkInside = 0x10,
            canLinkByOle1 = 0x20,
            isLinkObject = 0x40,
            insideOut = 0x80,
            activateWhenVisible = 0x100,
            renderingIsDeviceIndependent = 0x200,
            invisibleAtRuntime = 0x400,
            alwaysRun = 0x800,
            actsLikeButton = 0x1000,
            actsLikeLabel = 0x2000,
            noUiActivate = 0x4000,
            alignable = 0x8000,
            simpleFrame = 0x10000,
            setClientSiteFirst = 0x20000,
            imeMode = 0x40000,
            ignoreActivateWhenVisible = 0x80000,
            wantsToMenuMerge = 0x100000,
            supportsMultiLevelUndo = 0x200000
        }

        [FlagsAttribute]
        private enum TYPEFLAGS : short
        {
            FAPPOBJECT = 0x1,
            FCANCREATE = 0x2,
            FLICENSED = 0x4,
            FPREDECLID = 0x8,
            FHIDDEN = 0x10,
            FCONTROL = 0x20,
            FDUAL = 0x40,
            FNONEXTENSIBLE = 0x80,
            FOLEAUTOMATION = 0x100,
            FRESTRICTED = 0x200,
            FAGGREGATABLE = 0x400,
            FREPLACEABLE = 0x800,
            FDISPATCHABLE = 0x1000,
            FREVERSEBIND = 0x2000,
            FPROXY = 0x4000
        }

        /// <summary>
        /// Generates a manifest for the executable file (ActiveX .dll or .ocx)
        /// specified in <code>objectFilename</code>.
        /// </summary>
        /// 
        /// <param name="objectFilename">
        /// The path and Filename of the required executable file.
        /// </param>
        /// <returns></returns>
        public MemoryStream GenerateFromObjectFile(string objectFilename, string description = "")
        {
            var output = new MemoryStream();
            using (var w = XmlWriter.Create(output, new XmlWriterSettings() { Encoding = Encoding.UTF8, Indent = true, IndentChars = "    ", NewLineHandling=NewLineHandling.Entitize }))
            { 
                generateManifestXml(w,
                                    Path.GetFileNameWithoutExtension(objectFilename),
                                    fileVersionFromFileVersionInfo(FileVersionInfo.GetVersionInfo(objectFilename)),
                                    description,
                                    new Action[]
                                    {
                                        () => generateTypesInfo(objectFilename, w)
                                    }
                                    );
            }
            return output;
        }

        /// <summary>
        /// Generates a manifest for the Visual Basic 6  project
        /// specified in <code>projectFilename</code>.
        /// </summary>
        /// 
        /// <param name="projectFilename">
        /// The path and Filename of the required Visual Basic 6 project.
        /// </param>
        /// <param name="useVersion6CommonControls"
        /// Indicates whether version 6 of the Windows Common Controls are 
        /// to be used. 
        /// 
        /// If the program does not use the Windows Common 
        /// Controls then this parameter is ignored.
        /// <returns></returns>
        public MemoryStream GenerateFromProject(string projectFilename, bool useVersion6CommonControls)
        {
            string projectPath = Path.GetDirectoryName(projectFilename);
            string objectFilePath = String.Empty;
            string objectFilename = String.Empty;
            string version = String.Empty;
            string type = String.Empty;
            string description = String.Empty;

            var referenceLines = new List<string>();
            var objectLines = new List<string>();

            processProjectFile(projectFilename, referenceLines, objectLines, ref objectFilePath, ref objectFilename, ref type, ref version, ref description);
            if (! String.IsNullOrEmpty(projectPath)) objectFilePath = projectPath + @"\" + objectFilePath;
            
            var output = new MemoryStream();
            using (var w = XmlWriter.Create(output, new XmlWriterSettings() { Encoding = Encoding.UTF8, Indent = true, IndentChars = "    ", NewLineHandling = NewLineHandling.Entitize }))
            {
                var interfaces = new Dictionary<string, InterfaceInfo>();
                generateManifestXml(w,
                                    Path.GetFileNameWithoutExtension(objectFilename),
                                    version,
                                    description,
                                    new Action[]
                                    {
                                        () => referenceLines.ForEach(line => processReferenceLine(line, type, ref interfaces, w)),
                                        () => objectLines.ForEach(line => processObjectLine(line, type, ref interfaces, w, useVersion6CommonControls)),
                                        () => 
                                        {
                                            if (!type.Equals("Exe")) 
                                            {
                                                generateTypesInfo(objectFilePath + @"\" + objectFilename, w);
                                            }
                                            else 
                                            {
                                                if (useVersion6CommonControls) outputDependentAssembly(w, "Microsoft.Windows.Common-Controls", "6.0.0.0", "6595b64144ccf1df");
                                                generateComInterfaceExternalProxyStubElements(interfaces, w);
                                            }
                                        }
                                    });
            }
            return output;
        }

        private void processProjectFile(
            string projectFilename,
            List<string> referenceLines,
            List<string> objectLines,
            ref string objectFilePath,
            ref string objectFilename,
            ref string type,
            ref string version,
            ref string description)
        {
            int majorVersion = 0;
            int minorVersion = 0;
            int revisionVersion = 0;

            using (StreamReader sr = new StreamReader(projectFilename))
            {
                while (!sr.EndOfStream)
                {
                    processLine(sr.ReadLine(), referenceLines, objectLines, ref objectFilePath, ref objectFilename, ref type, ref majorVersion, ref minorVersion, ref revisionVersion, ref description);
                }
            }
            version = majorVersion.ToString() + "." + minorVersion.ToString() + ".0." + revisionVersion.ToString();
        }

        private void processLine(
            string line, 
            List<string> referenceLines, 
            List<string> objectLines, 
            ref string objectFilePath, 
            ref string objectFilename,
            ref string type,
            ref int majorVersion,
            ref int minorVersion,
            ref int revisionVersion,
            ref string description
            )
        {
            if (String.IsNullOrEmpty(line)) return;

            var index = line.IndexOf('=');
            if (index < 0) return;

            var lineType = line.Substring(0, index);
            var lineContent = line.Substring(index + 1, line.Length - index - 1);

            switch (lineType)
            {
            case "Description":
                description = Utils.trimDelimiters(lineContent);
                break;
            case "ExeName32":
                objectFilename = Utils.trimDelimiters(lineContent);
                break;
            case "MajorVer":
                majorVersion = Int32.Parse(lineContent);
                break;
            case "MinorVer":
                minorVersion = Int32.Parse(lineContent);
                break;
            case "Object":
                objectLines.Add(lineContent);
                break;
            case "Path32":
                objectFilePath =  Utils.trimDelimiters(lineContent);
                break;
            case "Reference":
                referenceLines.Add(lineContent);
                break;
            case "RevisionVer":
                revisionVersion = Int32.Parse(lineContent);
                break;
            case "Type":
                type = processTypeLine(lineContent);
                break;
            }
        }

        private static void processObjectLine(string lineContent, string projectType, ref Dictionary<string, InterfaceInfo> interfaces, XmlWriter w, bool useVersion6CommonControls)
        {
            string guid = getGuid(lineContent);
            string version = getObjectTypelibVersion(lineContent);
            string objectFilename = getObjectFilename(guid, version);

            generateDependentAssembly(objectFilename, w);
            if (projectType.Equals("Exe"))
            {
                extractExternalInterfaces(objectFilename, ref interfaces);
            }
        }

        private static void processReferenceLine(string lineContent, string projectType, ref Dictionary<string, InterfaceInfo> interfaces, XmlWriter w)
        {
            string guid = getGuid(lineContent);

            // ignore the following because they don't meed to be side-by-sided
            if (guid.Equals("{00020430-0000-0000-C000-000000000046}", StringComparison.CurrentCultureIgnoreCase)        /* stdole2 */
                || guid.Equals("{420B2830-E718-11CF-893D-00A0C9054228}", StringComparison.CurrentCultureIgnoreCase)     /* scrrun */
                || guid.Equals("{3F4DACA7-160D-11D2-A8E9-00104B365C9F}", StringComparison.CurrentCultureIgnoreCase)     /* vbscript */
                || guid.Equals("{F5078F18-C551-11D3-89B9-0000F81FE221}", StringComparison.CurrentCultureIgnoreCase)     /* msxml6 */
                || guid.Equals("{7C0FFAB0-CD84-11D0-949A-00A0C91110ED}", StringComparison.CurrentCultureIgnoreCase)     /* msdatsrc */
                || guid.Equals("{F5078F18-C551-11D3-89B9-0000F81FE221}", StringComparison.CurrentCultureIgnoreCase)     /* msxml6 */
                || guid.Equals("{2A75196C-D9EB-4129-B803-931327F72D5C}", StringComparison.CurrentCultureIgnoreCase)     /* msdao28 */
                )
            {
                return;
            }

            string referenceFilename = getReferenceFilename(guid);

            if (!referenceFilename.EndsWith(".tlb"))
            {
                generateDependentAssembly(referenceFilename, w);
                if (projectType.Equals("Exe"))
                {
                    extractExternalInterfaces(referenceFilename, ref interfaces);
                }
            }
        }

        private static string processTypeLine(string lineContent)
        {
            if (!lineContent.Equals("Exe") &&
                !lineContent.Equals("Control") &&
                !lineContent.Equals("OleDll")) throw new ArgumentException("Wrong project type");
            return lineContent;
        }

        private void generateManifestXml(XmlWriter w, string assemblyName, String version, string description, Action[] contentGenerators)
        {
            w.WriteStartDocument(true);
            w.WriteStartElement("assembly", "urn:schemas-microsoft-com:asm.v1");
                w.WriteAttributeString("manifestVersion", "1.0");
                w.WriteAttributeString("xmlns","asmv3",null, "urn:schemas-microsoft-com:asm.v3");

                w.WriteStartElement("assemblyIdentity");
                    w.WriteAttributeString("name", assemblyName);
                    w.WriteAttributeString("processorArchitecture", "X86");
                    w.WriteAttributeString("type", "win32");
                    w.WriteAttributeString("version", version);
                w.WriteEndElement();

                w.WriteElementString("description", description);

                foreach (Action f in contentGenerators) if (f != null ) f.Invoke();

            w.WriteEndElement();
            w.Flush();
        }

        private void generateTypesInfo(string assemblyFilename, XmlWriter w)
        {
            w.WriteStartElement("file");
            w.WriteAttributeString("name", Path.GetFileName(assemblyFilename));
            var tlia = new TLI.TLIApplication();
            var typelibInfo = tlia.TypeLibInfoFromFile(assemblyFilename);

            generateTypelibElement(typelibInfo, w);
            generateComClassElements(typelibInfo, w);
            w.WriteEndElement();
        }

        private static void generateTypelibElement(TLI.TypeLibInfo typelibInfo, XmlWriter w)
        {
            w.WriteStartElement("typelib");
            w.WriteAttributeString("tlbid", typelibInfo.GUID);
            w.WriteAttributeString("version", typelibInfo.MajorVersion.ToString() + "." + typelibInfo.MinorVersion.ToString());
            w.WriteAttributeString("flags", Enum.Format(typeof(LIBFLAGS), (LIBFLAGS)typelibInfo.AttributeMask, "G").Replace(", ",","));
            w.WriteAttributeString("helpdir", "");
            w.WriteEndElement();
        }

        private static void generateComClassElements(TypeLibInfo typelibInfo, XmlWriter w)
        {
            foreach (CoClassInfo c in typelibInfo.CoClasses)
            {
                if (!String.IsNullOrEmpty((string)Registry.GetValue(@"HKEY_CLASSES_ROOT\ClsId\" + c.GUID + @"\InprocServer32", "ThreadingModel", "")))
                {
                    generateComClassElement(typelibInfo, w, c);
                }
            }
        }

        private static void generateComClassElement(TypeLibInfo typelibInfo, XmlWriter w, CoClassInfo c)
        {
            w.WriteStartElement("comClass");
            w.WriteAttributeString("clsid", c.GUID);
            w.WriteAttributeString("tlbid", typelibInfo.GUID);

            string progID = null;
            if ((c.AttributeMask & (short)TYPEFLAGS.FHIDDEN) == 0)
            {
                progID = (string)Registry.GetValue(@"HKEY_CLASSES_ROOT\ClsId\" + c.GUID + @"\VersionIndependentProgID", "", "");
                if (String.IsNullOrEmpty(progID)) progID = (string)Registry.GetValue(@"HKEY_CLASSES_ROOT\ClsId\" + c.GUID + @"\ProgID", "", "");

                string curVer = (string)Registry.GetValue(@"HKEY_CLASSES_ROOT\" + progID + @"\CurVer", "", "");
                
                if (!String.IsNullOrEmpty(curVer))
                {
                    w.WriteAttributeString("progid", curVer);
                }
                else if (!String.IsNullOrEmpty(progID)) 
                {
                    w.WriteAttributeString("progid", progID);
                    progID = null;
                }
            }

            string threadingModel = (string)Registry.GetValue(@"HKEY_CLASSES_ROOT\ClsId\" + c.GUID + @"\InprocServer32", "ThreadingModel", "");
            if (!String.IsNullOrEmpty(threadingModel)) w.WriteAttributeString("threadingModel", threadingModel);

            string miscStatus = String.Empty;
            foreach (int i in Enum.GetValues(typeof(MISCSTATUS)))
            {
                int flags;
                Int32.TryParse((string)Registry.GetValue(@"HKEY_CLASSES_ROOT\ClsId\" + c.GUID + @"\MiscStatus" + (string)(i > 0 ? @"\" + i.ToString() : String.Empty), "", ""), out flags);
                if (flags != 0)
                {
                    w.WriteAttributeString(Enum.GetName(typeof(MISCSTATUS), i), Enum.Format(typeof(OLEMISC), flags, "G").Replace(", ",","));
                }
            }

            if (!String.IsNullOrEmpty(progID))
            {
                w.WriteElementString("progid", progID);
            }

            w.WriteEndElement();
        }

        private static void extractExternalInterfaces(string filename, ref Dictionary<string, InterfaceInfo> interfaces)
        {
            var tlia = new TLI.TLIApplication();
            var typelibInfo = tlia.TypeLibInfoFromFile(filename);
            foreach (CoClassInfo c in typelibInfo.CoClasses)
            {
                if (c.DefaultInterface != null)
                {
                    if (!interfaces.ContainsKey(c.DefaultInterface.GUID)) interfaces.Add(c.DefaultInterface.GUID, c.DefaultInterface);
                }
            }
        }

        private static void generateComInterfaceExternalProxyStubElements(Dictionary<string, InterfaceInfo> interfaces, XmlWriter w)
        {
            foreach (InterfaceInfo i in interfaces.Values)
            {
                generateComInterfaceExternalProxyStubElement(i, w);
            }
        }

        private static void generateComInterfaceExternalProxyStubElement(InterfaceInfo i, XmlWriter w)
        {
            w.WriteStartElement("comInterfaceExternalProxyStub");
            w.WriteAttributeString("name",i.Name);
            w.WriteAttributeString("iid", i.GUID);
            w.WriteAttributeString("proxyStubClsid32", (string)Registry.GetValue(@"HKEY_CLASSES_ROOT\Interface\" + i.GUID + @"\ProxyStubClsid32", "", ""));
            w.WriteEndElement();
        }

        private static void outputTypelibFileInfo(XmlWriter w, string guid, string objectFilename)
        {
            w.WriteStartElement("file");
            w.WriteAttributeString("name", Path.GetFileName(objectFilename));

            var tlia = new TLI.TLIApplication();
            var typelibInfo = tlia.TypeLibInfoFromFile(objectFilename);

            generateTypelibElement(typelibInfo,w);

            w.WriteEndElement();
        }

        private static void generateDependentAssembly(string objectFilename, XmlWriter w)
        {
            outputDependentAssembly(w, Path.GetFileNameWithoutExtension(objectFilename), fileVersionFromFileVersionInfo(FileVersionInfo.GetVersionInfo(objectFilename)), "");
        }

        private static void outputDependentAssembly(XmlWriter w, string assemblyName, String version, string publicKeyToken)
        {
            w.WriteStartElement("dependency");
            w.WriteStartElement("dependentAssembly");

            w.WriteStartElement("assemblyIdentity");
            w.WriteAttributeString("name", assemblyName);
            w.WriteAttributeString("processorArchitecture", "X86");
            w.WriteAttributeString("type", "win32");
            w.WriteAttributeString("version", version);
            if (publicKeyToken != String.Empty) w.WriteAttributeString("publicKeyToken", publicKeyToken);
            w.WriteEndElement();

            w.WriteEndElement();
            w.WriteEndElement();
        }

        private static string getObjectTypelibVersion(string line)
        {
            string pattern = @"\#([0-9]+)\.([0-9]+)\#";
            var match = Regex.Match(line, pattern);

            if (!match.Success) throw new InvalidOperationException(String.Format("No typelib version found in string {0}", line));
            return Utils.numericStringToHex(match.Groups[1].Value) + "." + Utils.numericStringToHex(match.Groups[2].Value);
        }

        private static string getObjectFilename(string guid, string version)
        {
            return (string)Registry.GetValue(@"HKEY_CLASSES_ROOT\Typelib\" + guid + @"\" + version + @"\0\win32", "", "");
        }

        private static string getReferenceFilename(string guid)
        {
            // Note that the Reference= lines in the project file may contain out-of-date
            // version information. So we look for the latest version
            string Filename = String.Empty;
            using (var regKey = Registry.ClassesRoot.OpenSubKey(@"Typelib\" + guid))
            {
                foreach (string subKeyName in regKey.GetSubKeyNames())
                {
                    using (var subKey = regKey.OpenSubKey(subKeyName + @"\0\win32"))
                    {
                        string s = (string)subKey.GetValue("");
                        if (s != String.Empty) Filename = s;
                    }
                }
            }

            if (Filename == String.Empty) throw new InvalidOperationException("Can't find Filename for guid " + guid);
            return Filename;
        }

        private static string getGuid(string line)
        {
            // include the surrounding braces
            string pattern = @"(?i)\{([0-9]|[A-F]){8}\-([0-9]|[A-F]){4}\-([0-9]|[A-F]){4}\-([0-9]|[A-F]){4}\-([0-9]|[A-F]){12}\}";
            var match = Regex.Match(line, pattern);

            if (!match.Success) throw new InvalidOperationException(String.Format("No GUID found in string {0}", line));
            return match.Groups[0].Value;
        }

        private static String fileVersionFromFileVersionInfo(FileVersionInfo versionInfo)
        {
            return versionInfo.FileMajorPart + "." + versionInfo.FileMinorPart + "." + versionInfo.FileBuildPart + "." + versionInfo.FilePrivatePart;
        }

    }
}
