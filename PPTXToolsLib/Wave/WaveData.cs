using System.Collections.Generic;
using System.Linq;

namespace PPTXTools.Wave
{
    /// <summary>
    /// 音声データを読み込む。
    /// このクラスでは読み込んっだ音声データから抽出されたチャンクをデリゲートによって通知する機能を提供する。
    /// 通知の際、音声データのサンプリングレート、bps、チャンネル数も通知する。
    /// チャンネル数は強制的に1とし、元の音声データ（チャンネル数2）をモノラル化してから通知する。
    /// </summary>
    public class WaveData
    {
        /// <summary>
        /// 音声データが抽出されたときに通知を受けるためのデリゲート
        /// </summary>
        /// <param name="channels">チャンネル数</param>
        /// <param name="samplingRate">サンプリングレート</param>
        /// <param name="bitPerSample">bps</param>
        /// <param name="data">チャンク本体</param>
        public delegate void WaveScanner(ushort channels, uint samplingRate, ushort bitPerSample, short[] data);

        private ushort _channels;
        private uint _samplingRate;
        private ushort _bitPerSample;

        private string _path;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="path">音声データのパス</param>
        public WaveData(string path)
        {
            _path = path;
        }

        /// <summary>
        /// 音声データの読み込みを開始する。
        /// 読み込み中はscanner による通知が順次行われる。
        /// </summary>
        /// <param name="scanner">通知を行うためのデリゲート</param>
        public void Scan(WaveScanner scanner)
        {
            ///List<short> data = new List<short>();
            using (WaveReaderS16Ch2 wr = new WaveReaderS16Ch2(_path))
            {
                ChunkFormat fmt = null;

                foreach (ChunkBase chunk in wr.ReadChunks())
                {
                    if (chunk is ChunkFormat)
                    {
                        fmt = (ChunkFormat)chunk;

                        _channels = fmt.Channels;
                        _samplingRate = fmt.SamplingRate;
                        _bitPerSample = fmt.BitPerSample;
                    }
                    else if (chunk is ChunkData)
                    {
                        foreach (short[,] wave in ((ChunkData)chunk).ReadDataS16x2((int)fmt.SamplingRate))
                        {
                            short[] crossSequence = wave.Cast<short>().ToArray();
                            short[] singleSequence = Enumerable.Range(0, crossSequence.Length / 2)
                                .Select(s => (short)(((int)crossSequence[s * 2] + (int)crossSequence[s * 2 + 1]) / 2)).ToArray();
                            scanner(1, _samplingRate, _bitPerSample, singleSequence); 
                        }
                    }
                }
            }
        }
    }
}
