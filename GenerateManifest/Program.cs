#region License

// The MIT License (MIT)
//
// Copyright (c) 2017-2018 Richard L King (TradeWright Software Systems)
// 
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
// 
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.
// 
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.

#endregion
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
            internal bool inlineExternalObjects;
            internal string assemblyName;
            internal string assemblyVersion;
            internal string assemblyDescription;
            internal IEnumerable<string> assemblyFiles;
            internal IEnumerable<string> dependentAssemblyIDs;
        }
        static int Main(string[] args)
        {
            try
            {
                var parameters = new Parameters();
                if (!GetParameters(ref parameters)) return 1;


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
            catch (ArgumentException e)
            {
                Console.WriteLine(e.Message);
                return 1;
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
            parameters.inlineExternalObjects = clp.get_IsSwitchSet("INLINE");

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
                var assClp = new CommandLineParser(parameters.assemblyInfo, ",");
                if (assClp.NumberOfArgs != 4)
                {
                    Console.WriteLine("Invalid arguments\n");
                    ShowUsage();
                    return false;
                }

                parameters.assemblyName = assClp.get_Arg(0);
                parameters.assemblyVersion = assClp.get_Arg(1);
                parameters.assemblyDescription = assClp.get_Arg(2);
                if (!File.Exists(assClp.get_Arg(3)))
                {
                    Console.WriteLine("Invalid argument: file {0} does not exist", assClp.get_Arg(3));
                    return false;
                }
                parameters.assemblyFiles = from f in LineReader(assClp.get_Arg(3))
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

                 IEnumerable<string> manFileReader = null;
                if (File.Exists(parameters.projectFile + ".man"))
                {
                    manFileReader = LineReader(parameters.projectFile + ".man");
                }

                IEnumerable<string> depFileReader = null;
                var depFilename = clp.get_SwitchValue("DEP");
                if (! string.IsNullOrEmpty(depFilename))
                {
                    if (!File.Exists(depFilename))
                    {
                        Console.WriteLine("Invalid argument: file {0} does not exist", depFilename);
                        return false;
                    }
                    depFileReader = LineReader(depFilename);
                }

                if (manFileReader != null && depFileReader != null)
                {
                    parameters.dependentAssemblyIDs = manFileReader.Union(depFileReader);
                }
                else if (manFileReader != null)
                {
                    parameters.dependentAssemblyIDs = manFileReader;
                }
                else if (depFileReader != null)
                {
                    parameters.dependentAssemblyIDs = depFileReader;
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
                data = gen.GenerateFromProject(parameters.projectFile, parameters.useVersion6CommonControls, parameters.inlineExternalObjects, parameters.dependentAssemblyIDs);
            }
            else if (!String.IsNullOrEmpty(parameters.objectFile))
            {
                data = gen.GenerateFromObjectFile(parameters.objectFile, parameters.description);
            }
            else if (!String.IsNullOrEmpty(parameters.assemblyName))
            {
                data = gen.GenerateFromFiles(parameters.assemblyName, parameters.assemblyVersion, parameters.assemblyDescription, parameters.inlineExternalObjects, parameters.assemblyFiles);
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
            Console.Write(
@"Creates a manifest for a dll or exe, or for a multifile assembly.
                 

GenerateManifest {/Proj:<projectFileName> [/V6CC] /Dep:<depFilename> | 
                  /Bin:<objectFilename> [/Desc:<description>] |
                  /Ass:<name>,<version>,<description>,<filenamesFile>}
                 [/Inline]
                 [/Out:<outputManifestFilename>]

    /Proj        Creates a manifest for the Visual Basic 6 project in 
                 <projectFileName>. If a file called <projectFileName>.man
                 exists or /Dep:<depFilename> is specified, then the manifest
                 is generated with the dependent assemblies specified in 
                 <projectFileName>.man and/or <depFilename>, instead of
                 taking the dependencies from the project file (see below 
                 for further details).

    /V6CC        Specifies that the project uses Microsoft Common Controls 
                 Version 6 (ignored if not an exe project).

    /Dep         Specifies a file containing external dependencies for this
                 project (see below for further details).

    /Inline      Specifies that <dependentAsssembly> elements are not to be
                 included for external references. Rather a <file> element
                 containing the COM class information is to be generated for
                 each external reference. Ignored if a <projectFileName>.man
                 file exists. This switch is relevant only when /Proj or /Ass
                 are specified, and is ignored otherwise.

    /Bin         Creates a manifest for the exe, dll or ocx in 
                 <objectFilename>.

    /Ass         Creates a multi-file assembly manifest for the projects or
                 object files contained in file <filenamesFile>.

    /Desc        This is used for the manifest's description element.

    /Out         Stores the manifest in <outputManifestFilename>. If not 
                 supplied, the manifest is written to stdout.

<projectFileName>.man and <depFilename>
                 Contains details of one dependent assembly per line. Each 
                 line is formatted as a complete <assemblyIdentity> element, 
                 including the final </assemblyIdentity> tag. Blank lines and 
                 lines beginning // are ignored.

<filenamesFile>
                 Contains one project or object filename per line. Object
                 files must be .dll or .ocx files. Blank lines and lines 
                 beginning // are ignored.
    
");
        }
    }
}
