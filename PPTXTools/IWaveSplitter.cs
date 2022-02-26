using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PPTXTools
{
    /// <summary>
    /// 音声データの有音区間を取得する
    /// </summary>
    public interface IWaveSplitter
    {
        /// <summary>
        /// 有音区間を検出する
        /// </summary>
        /// <param name="expectedLength">想定される有音区間の数</param>
        /// <param name="start">有音区間の対象となる音声の開始位置(秒)</param>
        /// <param name="end">有音区間の対象となる音声の終端位置(秒)</param>
        /// <returns>抽出された有音区間の数</returns>
        int Scan(int expectedLength, float start, float end);

        /// <summary>
        /// 抽出された有音区間を取得する
        /// </summary>
        /// <returns>(有音区間の開始位置（秒）, 有音区間の終端位置(秒))の配列</returns>
        (float, float)[] GetRangesSec();
    }
}
