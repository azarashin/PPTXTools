using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace PPTXTools.Wave
{
    /// <summary>
    /// 音声ファイルを読み込む。読み込むデータはWAV形式のデータとし、
    /// 16bit 2ch データのみを対象とする。
    /// このクラスでは、音声ファイルの属性を調べ、音声の実体部分をチャンクとして抽出する。
    /// </summary>
    public class WaveReaderS16Ch2 : IDisposable
    {
        private readonly FileStream _fs;
        private readonly UInt32 _fileSize;

        private long _chunkPos;
        private long _chunkSize;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="path">読み込み対象の音声ファイルへのパス</param>
        public WaveReaderS16Ch2(string path)
        {
            _fs = new FileStream(path, FileMode.Open, FileAccess.Read);

            if ("RIFF" != readStr(4, Encoding.ASCII))
            {
                throw (new Exception());
            }

            _fileSize = read32u();

            if ("WAVE" != readStr(4, Encoding.ASCII))
            {
                throw (new Exception());
            }

            _chunkPos = _fs.Position;
            _chunkSize = 0;
        }

        /// <summary>
        /// チャンクを読み込む
        /// </summary>
        /// <returns>読み込んだチャンクの列挙子</returns>
        public IEnumerable<ChunkBase> ReadChunks()
        {
            while (true)
            {
                byte[] tmp = new byte[8];

                _fs.Position = _chunkPos + _chunkSize;

                if (_fs.Read(tmp, 0, tmp.Length) != tmp.Length)
                {
                    yield break;
                }

                string type = Encoding.ASCII.GetString(tmp, 0, 4);
                UInt32 size = BitConverter.ToUInt32(tmp, 4);

                long chunkStart = _fs.Position;
                _chunkPos = chunkStart + size;

                switch (type)
                {
                    case "fmt ":
                        yield return (new ChunkFormat(type, size, _fs, chunkStart));
                        break;

                    case "data":
                        yield return (new ChunkData(type, size, _fs, chunkStart));
                        break;

                    default:
                        yield return (new ChunkBase(type, size, _fs, chunkStart));
                        break;
                }
            }
        }

        protected UInt32 read32u()
        {
            byte[] buff = new byte[4];
            _fs.Read(buff, 0, buff.Length);
            return (BitConverter.ToUInt32(buff, 0));
        }

        protected string readStr(int bytes, Encoding encoding)
        {
            byte[] buff = new byte[bytes];
            _fs.Read(buff, 0, buff.Length);
            return (encoding.GetString(buff));
        }

        #region IDisposable Support
        private bool _disposedValue = false; // 重複する呼び出しを検出する

        public void Dispose()
        {
            if (!_disposedValue)
            {
                _fs.Dispose();
                _disposedValue = true;
            }
        }
        #endregion

    }
}
