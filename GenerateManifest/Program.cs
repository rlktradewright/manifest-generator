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
        static void Main(string[] args)
        {
            CommandLineParser clp = new CommandLineParser(Environment.CommandLine, " ");

            if (clp.get_IsSwitchSet("?") || clp.get_IsSwitchSet("HELP"))
            {
                ShowUsage();
                return;
            }
            
            string projectFile = clp.get_SwitchValue("PROJ");
            string objectFile = clp.get_SwitchValue("BIN");
            string description = clp.get_SwitchValue("DESC");
            string outFile = clp.get_SwitchValue("OUT");

            var gen = new ManifestGenerator();
            MemoryStream data;

            if (projectFile != String.Empty)
            {
                data = gen.GenerateFromProject(projectFile, true);
            }
            else if (objectFile != String.Empty)
            {
                data = gen.GenerateFromObjectFile(objectFile, description);
            }
            else
            {
                Console.WriteLine("Invalid arguments\n");
                ShowUsage();
                return;
            }

            if (outFile != String.Empty)
            {
                using (FileStream fs = new FileStream(outFile, FileMode.Create))
                {
                    data.WriteTo(fs);
                }
            }
            else
            {
                data.Position = 0;

                XmlReader reader = XmlReader.Create(data);
                reader.Read();
                Console.WriteLine("<?xml {0} ?>", reader.Value);
                reader.MoveToContent();
                Console.WriteLine(reader.ReadOuterXml());
                reader.Close();
            }

            data.Dispose();
        }

        private static void ShowUsage()
        {
            Console.Write(@"Creates a manifest for a dll or exe

GenerateManifest {/Proj:<projectFileName> | /Bin:<exeOrDLlName>} 
                 [/Desc:<description>] 
                 [/Out:<outputManifestFilename>]


");
        }
    }
}
