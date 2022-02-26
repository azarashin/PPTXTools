using System;
using System.IO;

namespace PPTXTools.Wave
{
    /// <summary>
    /// チャンクデータの共通部分
    /// </summary>
    public class ChunkBase
    {
        private readonly string _type;
        /// <summary>
        /// チャンク種別
        /// </summary>
        public string Type { get { return (_type); } }

        private readonly UInt32 _size;
        /// <summary>
        /// チャンクサイズ
        /// </summary>
        public UInt32 Size { get { return (_size); } }

        protected readonly FileStream _fs;
        protected readonly long _chunkStart; // chunk head + 8 (type + size)

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="type">チャンク種別。</param>
        /// <param name="size">チャンクサイズ</param>
        /// <param name="fs">読み込み元のファイルストリーム</param>
        /// <param name="chunkStart">ファイルストリーム上の読み込み位置</param>
        public ChunkBase(string type, UInt32 size, FileStream fs, long chunkStart)
        {
            _type = type;
            _size = size;
            this._fs = fs;
            this._chunkStart = chunkStart;
        }

        protected void Seek(long a)
        {
            _fs.Position = _chunkStart + a;
        }

        protected int TryRead(out byte[] data, int length)
        {
            data = new byte[length];
            int result = _fs.Read(data, 0, length);
            return result;
        }
    }
}
