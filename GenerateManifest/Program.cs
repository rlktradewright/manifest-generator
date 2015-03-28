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
            var gen = new ManifestGenerator();
            MemoryStream data;

            string projectFile =String.Empty;
            string objectFile = String.Empty;
            string description = String.Empty;
            string outFile = String.Empty;

            foreach (string arg in args)
            {
                if (arg.ToUpper().StartsWith("/PROJ:"))
                {
                    projectFile = arg.Substring("/PROJ:".Length);
                }
                else if (arg.ToUpper().StartsWith("/BIN:"))
                {
                    objectFile = arg.Substring("/BIN:".Length);
                }
                else if (arg.ToUpper().StartsWith("/DESC:"))
                {
                    description = arg.Substring("/DESC:".Length);
                }
                else if (arg.ToUpper().StartsWith("/OUT:"))
                {
                    outFile = arg.Substring("/OUT:".Length);
                }
            }

            if (projectFile != String.Empty)
            {
                data =gen.GenerateFromProject(projectFile, true);
            }
            else if (objectFile != String.Empty)
            {
                data = gen.GenerateFromObjectFile(objectFile, description);
            }
            else
            {
                Console.WriteLine("Invalid arguments");
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
    }
}
