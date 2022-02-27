using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PPTXTools
{
    public class NoRecordSlideException : Exception
    {
        private int _page; 
        public NoRecordSlideException(int page)
        {
            _page = page; 
        }

        public override string ToString()
        {
            return $"{base.ToString()}\nスライド番号{_page}にスライドの記録がありません。";
        }
    }
}
