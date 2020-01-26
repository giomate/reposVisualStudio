using System;
using System.IO;
using System.IO.Compression;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ZipTest
{
    class Program
    {
        static void Main(string[] args)
        {
            string CD = Directory.GetCurrentDirectory();
            string startPath = Path.Combine(CD, "ars_2400024_all");
            string zipPath = Path.Combine(CD, "ars_2400024.zip");
            ZipFile.CreateFromDirectory(startPath, zipPath);


        }

    }

}
