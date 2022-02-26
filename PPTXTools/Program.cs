using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PPTXTools
{
    class Program
    {
        static void Main(string[] args)
        {
            string wavPath = "./_tmp.wav";
            string pptxPath;
            string mp4Path;
            string srtPath;
            PPTXLoader loader;
            IWaveSplitter ana;
            SRTGenerator gen; 

            Console.WriteLine("*** パワーポイント(.pptx)のファイルのパスを入力してください");
            pptxPath = Console.ReadLine().Trim();

            Console.WriteLine("*** パワーポイントから生成された動画ファイル(.mp4)のファイルのパスを入力してください");
            mp4Path = Console.ReadLine().Trim();

            Console.WriteLine("*** 出力先の字幕ファイル(.srt)のパスを入力してください");
            srtPath = Console.ReadLine().Trim();

            Console.WriteLine("... mp4 を音声データ(.wav)に変換しています...");
            try
            {
                Mp4ToWav.ConvertMp4ToWav(mp4Path, wavPath);
            } catch(Exception ex)
            {
                Console.WriteLine("! 音声データへの変換処理に失敗しました");
                Console.WriteLine(ex.ToString());
                return; 
            }
            Console.WriteLine("... 音声データ(.wav)への変換処理が完了しました...");

            Console.WriteLine("... パワーポイントを読み込んでいます...");
            try
            {
                loader = new PPTXLoader(pptxPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("! パワーポイントの読み込みに失敗しました");
                Console.WriteLine(ex.ToString());
                return;
            }
            Console.WriteLine("... パワーポイントを読み込みました...");

            //IWaveAnalyzer ana = new WaveAnalyzer(wav_path);
            Console.WriteLine("... 音声ファイルを読み込んでいます...");
            try
            {
                ana = new ParallelWaveSplitter(wavPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("! 音声ファイルの読み込みに失敗しました");
                Console.WriteLine(ex.ToString());
                return;
            }
            Console.WriteLine("... 音声ファイルを読み込みました...");

            Console.WriteLine("... 字幕ファイルを生成しています...");
            try
            {
                gen = new SRTGenerator(loader, ana);
            }
            catch (Exception ex)
            {
                Console.WriteLine("! 字幕の生成に失敗しました");
                Console.WriteLine(ex.ToString());
                return;
            }
            Console.WriteLine("... 字幕ファイルを生成しました...");

            Console.WriteLine("... 字幕ファイルを保存しています...");
            try
            {
                gen.DumpSRT(srtPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("! 字幕の保存に失敗しました");
                Console.WriteLine(ex.ToString());
                return;
            }
            Console.WriteLine("... 字幕ファイルを保存しました...");
        }
    }
}
