using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WordPadNecdetElbruz
{
    class Document
    {
        public bool IsEdit { get; set; } = false;
        public string PathDocument { get; set; } = null;

        public string GetNameDoc()
        {
            if (PathDocument != String.Empty)
            {
                return Path.GetFileNameWithoutExtension(PathDocument);
            }
            return String.Empty;
        }
    }
}
