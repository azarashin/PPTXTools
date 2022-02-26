using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PPTXTools
{
    /// <summary>
    /// FFMPEG を外部プログラムとして使用し、動画ファイルからWAVE データを抽出する
    /// </summary>
    public static class Mp4ToWav
    {
        /// <summary>
        /// FFMPEG を外部プログラムとして使用し、動画ファイルからWAVE データを抽出する
        /// </summary>
        /// <param name="path">動画ファイルのパス名</param>
        /// <param name="outPath">出力先音声ファイルのパス名</param>
        /// <returns></returns>
        public static void ConvertMp4ToWav(string path, string outPath)
        {
            string out_path = "./_tmp.wav";
            if(File.Exists(out_path))
            {
                File.Delete(out_path); 
            }
            Process p = Process.Start("ffmpeg", $"-i \"{path}\" \"{outPath}\"");
            p.WaitForExit(); 
        }
    }
}
