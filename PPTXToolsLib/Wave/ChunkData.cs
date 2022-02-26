using System;
using System.Collections.Generic;
using System.IO;

namespace PPTXTools.Wave
{
    public class ChunkData : ChunkBase
    {
        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="type">チャンク種別。</param>
        /// <param name="size">チャンクサイズ</param>
        /// <param name="fs">読み込み元のファイルストリーム</param>
        /// <param name="chunkStart">ファイルストリーム上の読み込み位置</param>
        public ChunkData(string type, UInt32 size, FileStream fs, long ChunkStart)
            : base(type, size, fs, ChunkStart)
        {

        }

        /// <summary>
        /// 指定されたサイズのデータを読み込む
        /// </summary>
        /// <param name="blockSize">読み込みサイズ</param>
        /// <returns>読み込まれたデータ[時系列成分,チャンネル成分]</returns>
        public IEnumerable<Int16[,]> ReadDataS16x2(int blockSize)
        {
            for (long dataPos = 0; dataPos < Size; dataPos += 2 * 2 * blockSize)
            {
                byte[] rawData;

                Seek(dataPos);
                int length = TryRead(out rawData, 2 * 2 * blockSize);
                if (length == 0 || length % (2 * 2) != 0)
                {
                    yield break;
                }
                length = length / (2 * 2);

                Int16[,] waveData = new Int16[length, 2];

                for (int i = 0; i < length; i++)
                {
                    waveData[i, 0] = BitConverter.ToInt16(rawData, i * 2 * 2 + 0);
                    waveData[i, 1] = BitConverter.ToInt16(rawData, i * 2 * 2 + 2);
                }

                yield return (waveData);
            }
        }
    }
}
