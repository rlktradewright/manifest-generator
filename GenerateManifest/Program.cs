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
            internal IEnumerable<string> dependentAssemblyIDs;
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
                       WriteManifestToFile(parameters, data);
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

            if (!getProjectFile(ref parameters, clp))
            {
                return false;
            } 
            else if (!getObjectFile(ref parameters, clp))
            {
                return false;
            } 
            else if (!getAssemblyDetails(ref parameters, ref clp))
            {
                return false;
            } 
        
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

            return true;
        }

        private static bool getAssemblyDetails(ref Parameters parameters, ref CommandLineParser clp)
        {
            if (clp.get_IsSwitchSet("ASS"))
            {
                parameters.assemblyInfo = clp.get_SwitchValue("ASS");
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
                if (!File.Exists(clp.get_Arg(3)))
                {
                    Console.WriteLine("Invalid argument: file {0} does not exist", clp.get_Arg(3));
                    return false;
                }
                parameters.assemblyProjectFiles = from f in LineReader(clp.get_Arg(3))
                                                  where !String.IsNullOrEmpty(f) && !f.StartsWith("//")
                                                  select f;
            }
            return true;
        }

        private static bool getObjectFile(ref Parameters parameters, CommandLineParser clp)
        {
            if (clp.get_IsSwitchSet("BIN"))
            {
                parameters.objectFile = clp.get_SwitchValue("BIN");
                if (!File.Exists(parameters.objectFile))
                {
                    Console.WriteLine("Invalid argument: file {0} does not exist", parameters.projectFile);
                    return false;
                }
            }
            return true;
        }

        private static bool getProjectFile(ref Parameters parameters, CommandLineParser clp)
        {
            if (clp.get_IsSwitchSet("PROJ"))
            {
                parameters.projectFile = clp.get_SwitchValue("PROJ");
                if (!File.Exists(parameters.projectFile))
                {
                    Console.WriteLine("Invalid argument: file {0} does not exist", parameters.projectFile);
                    return false;
                }
                if (File.Exists(parameters.projectFile + ".man"))
                {
                    parameters.dependentAssemblyIDs = LineReader(parameters.projectFile + ".man");
                }
            }
            return true;
        }

        private static MemoryStream GenerateManifest(Parameters parameters)
        {
            var gen = new ManifestGenerator();
            MemoryStream data = null;
            if (!String.IsNullOrEmpty(parameters.projectFile))
            {
                data = gen.GenerateFromProject(parameters.projectFile, parameters.useVersion6CommonControls, parameters.dependentAssemblyIDs);
            }
            else if (!String.IsNullOrEmpty(parameters.objectFile))
            {
                data = gen.GenerateFromObjectFile(parameters.objectFile, parameters.description);
            }
            else if (!String.IsNullOrEmpty(parameters.assemblyName))
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

        private static void WriteManifestToFile(Parameters parameters, MemoryStream data)
        {
            using (FileStream fs = new FileStream(parameters.outFile, FileMode.Create))
            {
                data.WriteTo(fs);
            }
        }

        private static void ShowUsage()
        {
//========1=========2=========3=========4=========5=========6=========7=========8
            Console.Write(@"Creates a manifest for a dll or exe, or for a multifile assembly.
                 

GenerateManifest {/Proj:<projectFileName> [/V6CC] | 
                  /Bin:<objectFilename> [/Desc:<description>] |
                  /Ass:<name>,<version>,<description>,<projectsFileName>}
                 [/Out:<outputManifestFilename>]

    /Proj        Creates a manifest for the Visual Basic 6 project in 
                 <projectFileName>. If a file called <projectFileName>.man
                 exists, then the manifest is generated with the dependent 
                 assemblies specified in <projectFileName>.man, instead of
                 taking the dependencies from the project file (see below 
                 for further details).
    /V6CC        Specifies that the project uses Microsoft Common Controls 
                 Version 6 (ignored if not an exe project).
    /Bin         Creates a manifest for the exe, dll or ocx in <objectFilename>.
    /Ass         Creates a multi-file assembly manifest for the projects 
                 contained in file <projectsFilename>.
    /Desc        This is used for the manifest's description element.
    /Out         Stores the manifest in <outputManifestFilename>. If not 
                 supplied, the manifest is written to stdout.

<projectFileName>.man 
                 Contains details of one dependent assembly per line.
                 Blank lines and lines beginning // are ignored. Each
                 line is formatted as a complete <assemblyIdentity>
                 element, including the final </assemblyIdentity> tag.

<projectsFilename> 
                 Contains one project filename per line. Blank lines
                 and lines beginning // are ignored.
    
");
        }
    }
}
