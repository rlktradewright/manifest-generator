using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using TradeWright.ManifestUtilities;

namespace TradeWright.GenerateManifest
{
    class Program
    {

        private struct Parameters
        {
            internal string projectFile;
            internal string objectFile;
            internal string assemblyInfo;
            internal string description;
            internal string outFile;
            internal bool useVersion6CommonControls;
            internal string assemblyName;
            internal string assemblyVersion;
            internal string assemblyDescription;
            internal IEnumerable<string> assemblyProjectFiles;
        }
        static int Main(string[] args)
        {
            try
            {
                var parameters = new Parameters();
                if (! GetParameters(ref parameters)) return 1;


                using (MemoryStream data = GenerateManifest(parameters))
                {

                    if (parameters.outFile != String.Empty)
                    {
                        parameters = WriteManifestToFile(parameters, data);
                    }
                    else
                    {
                        WriteManifestToConsole(data);
                    }

                    return 0;
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                return 1;
            }
        }

        private static bool GetParameters(ref Parameters parameters)
        {
            var clp = new CommandLineParser(Environment.CommandLine, " ");

            if (clp.get_IsSwitchSet("?") || clp.get_IsSwitchSet("HELP"))
            {
                ShowUsage();
                return false;
            }

            parameters.projectFile = clp.get_SwitchValue("PROJ");
            parameters.objectFile = clp.get_SwitchValue("BIN");
            parameters.assemblyInfo = clp.get_SwitchValue("ASS");
            parameters.description = clp.get_SwitchValue("DESC");
            parameters.outFile = clp.get_SwitchValue("OUT");
            parameters.useVersion6CommonControls = clp.get_IsSwitchSet("V6CC");

            if ((String.IsNullOrEmpty(parameters.projectFile) && String.IsNullOrEmpty(parameters.objectFile) && String.IsNullOrEmpty(parameters.assemblyInfo)) ||
                (!String.IsNullOrEmpty(parameters.projectFile) && !String.IsNullOrEmpty(parameters.objectFile) && !String.IsNullOrEmpty(parameters.assemblyInfo)))
            {
                Console.WriteLine("Invalid arguments: only one of /Proj, /Bin and /Ass may be supplied\n");
                ShowUsage();
                return false;
            }

            if (!String.IsNullOrEmpty(parameters.assemblyInfo))
            {
                clp = new CommandLineParser(parameters.assemblyInfo, ",");
                if (clp.NumberOfArgs != 4)
                {
                    Console.WriteLine("Invalid arguments\n");
                    ShowUsage();
                    return false;
                }

                parameters.assemblyName = clp.get_Arg(0);
                parameters.assemblyVersion = clp.get_Arg(1);
                parameters.assemblyDescription = clp.get_Arg(2);
                parameters.assemblyProjectFiles = from f in LineReader(clp.get_Arg(3))
                                                  where !String.IsNullOrEmpty(f) && !f.StartsWith("//")
                                                  select f;
            }

            return true;
        }

        private static MemoryStream GenerateManifest(Parameters parameters)
        {
            var gen = new ManifestGenerator();
            MemoryStream data = null;
            if (parameters.projectFile != String.Empty)
            {
                data = gen.GenerateFromProject(parameters.projectFile, parameters.useVersion6CommonControls);
            }
            else if (parameters.objectFile != String.Empty)
            {
                data = gen.GenerateFromObjectFile(parameters.objectFile, parameters.description);
            }
            else if (parameters.assemblyName != String.Empty)
            {
                data = gen.GenerateFromProjects(parameters.assemblyName, parameters.assemblyVersion, parameters.assemblyDescription, parameters.assemblyProjectFiles);
            }
            return data;
        }

        private static IEnumerable<String> LineReader(String fileName)
        {
            String line;
            using (var file = File.OpenText(fileName))
            {
                while ((line = file.ReadLine()) != null)
                {
                    yield return line.Trim();
                }
            }
        }
        
        private static void WriteManifestToConsole(MemoryStream data)
        {
            data.Position = 0;

            XmlReader reader = XmlReader.Create(data);
            reader.Read();
            Console.WriteLine("<?xml {0} ?>", reader.Value);
            reader.MoveToContent();
            Console.WriteLine(reader.ReadOuterXml());
            reader.Close();
        }

        private static Parameters WriteManifestToFile(Parameters parameters, MemoryStream data)
        {
            using (FileStream fs = new FileStream(parameters.outFile, FileMode.Create))
            {
                data.WriteTo(fs);
            }
            return parameters;
        }

        private static void ShowUsage()
        {
            Console.Write(@"Creates a manifest for a dll or exe

GenerateManifest {/Proj:<projectFileName> [/V6CC] | 
                  /Bin:<objectFilename> [/Desc:<description>] |
                  /Ass:<name>,<version>,<description>,<projectsFileName>}
                 [/Out:<outputManifestFilename>]

    /Proj        create manifest for project in <projectFileName>
    /V6CC        project uses Microsoft Common Controls Version 6 (ignored
                 if not an exe project)
    /Bin         create manifest for exe, dll or ocx in <objectFilename>
    /Ass         create multi-file assembly manifest for the projects 
                 contained in file <projectsFilename>
    /Desc        used for the manifest description
    /Out         store the manifest in <outputManifestFilename>

<projectsFilename> contains one project filename per line. Blank lines
and lines beginning // are ignored.
    
");
        }
    }
}
