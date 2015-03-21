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
            Console.WriteLine(gen.Generate(args[0], true));
        }
    }
}
