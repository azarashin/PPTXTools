using PPTXTools.Python;
using System;

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
//            pptxPath = @"F:\MyDevelopment\PPTXTools\doc\Demo.pptx";

            Console.WriteLine("*** パワーポイントから生成された動画ファイル(.mp4)のファイルのパスを入力してください");
            mp4Path = Console.ReadLine().Trim();
//            mp4Path = @"F:\MyDevelopment\PPTXTools\doc\Demo.mp4";

            Console.WriteLine("*** 出力先の字幕ファイル(.srt)のパスを入力してください");
            srtPath = Console.ReadLine().Trim();
//            srtPath = @"F:\MyDevelopment\PPTXTools\doc\Demo.srt";

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

            Console.WriteLine("... パワーポイントのエラーチェックを開始します...");
            PPTXErrorChecker pptxErrorChecker = new PPTXErrorChecker(pptxPath, wavPath); 
            Console.WriteLine(pptxErrorChecker);
            Console.WriteLine("... パワーポイントのエラーチェックを完了しました...");

            Console.WriteLine("... 音声ファイルを読み込んでいます...");
            try
            {
                ana = new PytohnInaSpeechSegmenter(wavPath);
            }
            catch (PytohnInaSpeechSegmenterException ex)
            {
                Console.WriteLine(ex);
                Console.WriteLine("pytohn のInaSpeechSegmenter の実行に失敗しました。");
                Console.WriteLine("簡易式の音区間検出アルゴリズムに切り替えます。");
                try
                {
                    // ana = new WaveAnalyzer(wav_path);
                    ana = new ParallelWaveSplitter(wavPath);
                }
                catch (Exception ex2)
                {
                    Console.WriteLine("! 音声ファイルの読み込みに失敗しました");
                    Console.WriteLine(ex2.ToString());
                    return;
                }
            }
            Console.WriteLine("... 音声ファイルを読み込みました...");

            Console.WriteLine("... 字幕ファイルを生成しています...");
            try
            {
                try
                {
                    loader.Adjust(mp4Path, pptxErrorChecker.TotalWaveLength()); 
                    gen = new SRTGenerator(loader, ana);
                }
                catch (NoRecordSlideException ex)
                {
                    Console.WriteLine(ex);
                    return; 
                }
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
                gen.DumpSRT(srtPath, mp4Path, pptxErrorChecker.TotalWaveLength());
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
