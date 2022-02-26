using System;
using System.IO;

namespace PPTXTools.Wave
{
    public class ChunkFormat : ChunkBase
    {
        private readonly UInt16 _channels;
        /// <summary>
        /// チャンネル数
        /// </summary>
        public UInt16 Channels { get { return (_channels); } }

        private readonly UInt32 _samplingRate;
        /// <summary>
        /// サンプリングレート
        /// </summary>
        public UInt32 SamplingRate { get { return (_samplingRate); } }

        private readonly UInt16 _bitPerSample;
        /// <summary>
        /// 量子ビット数
        /// </summary>
        public UInt16 BitPerSample { get { return (_bitPerSample); } }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="type">チャンク種別。</param>
        /// <param name="size">チャンクサイズ</param>
        /// <param name="fs">読み込み元のファイルストリーム</param>
        /// <param name="chunkStart">ファイルストリーム上の読み込み位置</param>
        public ChunkFormat(string Type, UInt32 Size, FileStream fs, long ChunkStart)
            : base(Type, Size, fs, ChunkStart)
        {
            Seek(0);

            bool success = false;

            byte[] data;
            if (false) { }
            else if (Size == 16 && TryRead(out data, (int)Size) == Size)
            {
                UInt16 formatId = BitConverter.ToUInt16(data, 0);

                if (formatId == 1)
                {
                    _channels = BitConverter.ToUInt16(data, 2);
                    _samplingRate = BitConverter.ToUInt32(data, 4);
                    _bitPerSample = BitConverter.ToUInt16(data, 14);

                    success = true;
                }
            }

            if (!success)
            {
                throw new Exception();
            }
        }
    }
}
