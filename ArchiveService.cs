using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsApp1
{
    internal class ArchiveService
    {
        private string _pathTo7z = string.Empty;
        private static readonly string command = string.Empty;
        internal ArchiveService(bool flag, string path)
        {
            _pathTo7z = path;
        }


    }
}
