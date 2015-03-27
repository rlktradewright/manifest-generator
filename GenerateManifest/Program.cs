using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using TradeWright.ManifestUtilities;

namespace TradeWright.GenerateManifest
{
    class Program
    {
        static void Main(string[] args)
        {
            var gen = new ManifestGenerator();
            if (args[0].ToUpper().StartsWith("/P"))
            {
                Console.WriteLine(gen.GenerateFromProject(args[1], true));
            }
            else if (args[0].ToUpper().StartsWith("/B"))
            {
                Console.WriteLine(gen.GenerateFromObjectFile(args[1]));
            }
            else
            {
                Console.WriteLine("Invalid arguments");
            }
            
        }
    }
}
