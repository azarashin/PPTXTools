using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PPTXTools.Python
{
    public class PytohnInaSpeechSegmenter : IWaveSplitter
    {
        private (float, float)[] _all;
        private (float, float)[] _current;

        public PytohnInaSpeechSegmenter(string pathToWavFile)
        {
            string outPath = "./_tmp.ina_speech_segment.txt";
            if (File.Exists(outPath))
            {
                File.Delete(outPath);
            }
            try
            {
                Process p = Process.Start("ina_speech_segmenter.bat", $"\"{pathToWavFile}\" \"{outPath}\"");
                p.WaitForExit();
                using (StreamReader sr = new StreamReader(outPath))
                {
                    _all = sr.ReadToEnd()
                        .Split('\n')
                        .Select(s => s.Split(','))
                        .Where(s => s[0] == "male" || s[0] == "female")
                        .Select(s => (float.Parse(s[1]), float.Parse(s[2])))
                        .ToArray();
                }
            }
            catch (Exception)
            {
                throw new PytohnInaSpeechSegmenterException();
            }
        }

        public (float, float)[] GetRangesSec()
        {
            return _current; 
        }

        public int Scan(int expectedLength, float start, float end)
        {
            _current = _all
                .Where(s => s.Item1 >= start && s.Item1 <= end)
                .ToArray();
            return _current.Length; 
        }
    }
}
