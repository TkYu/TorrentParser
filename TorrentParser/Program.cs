using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TorrentParser
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 0)
            {
                Torrent t;
                if(Torrent.TryParse(args[0],out t))
                    Console.WriteLine(t.MagnetURI);
                Console.ReadKey();
            }
        }
    }
}
