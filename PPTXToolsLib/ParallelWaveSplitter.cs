using PPTXTools.Wave;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PPTXTools
{
    /// <summary>
    /// 音声データの有音区間を取得する。
    /// 本クラスでは有音区間の数が想定される有音区間の数をなるべく超えるよう
    /// 内部パラメータが自動的に調整される。
    /// 有音区間の抽出結果と内部パラメータはキャッシュされるため、繰り返し実行する場合は
    /// 軽量である。
    /// また、キャッシュを生成するための処理が最適化されており、キャッシュの生成にかかる時間が比較的短い。
    /// 但し、クラスの構成は複雑であり、メンテナンスコストが比較的高い。
    /// </summary>
    public class ParallelWaveSplitter : IWaveSplitter
    {
        public (float, float)[] RangesSec { get; private set; }

        private string _path;

        List<ParallelWaveAnalyzerSub> _subAnalyzers;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="path">有音区間を抽出する対象の音声ファイルのパス</param>
        public ParallelWaveSplitter(string path)
        {
            _path = path;

            float speed = 0.05f;
            float bias = 0.001f;

            WaveData data = new WaveData(_path);
            WaveAttribute attr = new WaveAttribute(data);

            _subAnalyzers = new List<ParallelWaveAnalyzerSub>();
            for (int i = 10; i >= 0; i--)
            {
                float silentWeight = speed * i + bias;
                attr.UpdateSilentParameter(silentWeight);
                _subAnalyzers.Add(new ParallelWaveAnalyzerSub(_path, attr.SilentParameter));

            }

            data.Scan(ScanSplitParameter);
            foreach (var sub in _subAnalyzers)
            {
                sub.ScanPostProcess();

            }
        }

        /// <summary>
        /// 有音区間を検出する
        /// </summary>
        /// <param name="expectedLength">想定される有音区間の数</param>
        /// <param name="start">有音区間の対象となる音声の開始位置(秒)</param>
        /// <param name="end">有音区間の対象となる音声の終端位置(秒)</param>
        /// <returns>抽出された有音区間の数</returns>
        public int Scan(int expectedLength, float start, float end)
        {
            RangesSec = null; 
            foreach (var sub in _subAnalyzers)
            {
                int targetCount = sub.RangesSec
                    .Where(s => s.Item1 >= start && s.Item2 <= end)
                    .Count();
                if(targetCount >= expectedLength)
                {
                    RangesSec = sub.RangesSec;
                    break;
                }
            }

            return RangesSec.Length; 
        }

        private void ScanSplitParameter(ushort channels, uint samplingRate, ushort bitPerSample, short[] data)
        {
            foreach (var sub in _subAnalyzers)
            {
                sub.ScanSplitParameter(channels, samplingRate, bitPerSample, data);
            }
        }

        /// <summary>
        /// 抽出された有音区間を取得する
        /// </summary>
        /// <returns>(有音区間の開始位置（秒）, 有音区間の終端位置(秒))の配列</returns>
        public (float, float)[] GetRangesSec()
        {
            return RangesSec;
        }



        private class ParallelWaveAnalyzerSub
        {
            public const float MinDurationSec = 1.0f;

            private string _path;

            private float _silentThrethold;
            private int _keep = 0;
            private bool _mode = false; // true: 有音, falsle: 無音
            private int _count = 0;
            private int _start = -1;
            private List<(int, int)> _ranges;
            private uint _samplingRate;

            internal (float, float)[] RangesSec { get; private set; }

            /// <summary>
            /// コンストラクタ
            /// </summary>
            /// <param name="path">区間抽出対象の音声データのパス</param>
            /// <param name="silentThreshold">無音区間抽出パラメータ</param>
            internal ParallelWaveAnalyzerSub(string path, float silentThreshold)
            {
                _path = path;
                _silentThrethold = silentThreshold;

                _keep = 0;
                _mode = false;
                _count = 0;
                _start = -1;
                _ranges = new List<(int, int)>();
            }

            /// <summary>
            /// 音声データのスキャンが一通り終わった後に実行する処理
            /// </summary>
            /// <returns>検出された有音区間の数</returns>
            internal int ScanPostProcess()
            {
                (int, int)[] ranges;

                _ranges = _ranges.Where(s => s.Item2 - s.Item1 > MinDurationSec * _samplingRate).ToList();

                if (_ranges.Count() == 0)
                {
                    ranges = new (int, int)[0];
                }
                else
                {
                    List<(int, int)> range_final = new List<(int, int)>();
                    range_final.Add(_ranges[0]);
                    int threshold_sec = 1; // この秒数未満の区間は一つにまとめる
                    int threshold = threshold_sec * (int)_samplingRate;
                    for (int i = 1; i < _ranges.Count(); i++)
                    {
                        (int, int) lst = range_final.Last();
                        (int, int) cur = _ranges[i];
                        if (cur.Item2 - lst.Item1 < threshold)
                        {
                            range_final.RemoveAt(range_final.Count() - 1);
                            range_final.Add((lst.Item1, cur.Item2));
                        }
                        else
                        {
                            range_final.Add(cur);
                        }
                    }

                    ranges = range_final.ToArray();
                }
                RangesSec = ranges.Select(s => (s.Item1 / (float)_samplingRate, s.Item2 / (float)_samplingRate)).ToArray();

                return RangesSec.Length;
            }

            /// <summary>
            /// 音声データのチャンクに対して順次分析処理を実行する
            /// </summary>
            /// <param name="channels">チャンネル数。本メソッドでは1 の時のみ動作する。</param>
            /// <param name="samplingRate">サンプリングレート</param>
            /// <param name="bitPerSample">bps</param>
            /// <param name="data">チャンクデータ</param>
            internal void ScanSplitParameter(ushort channels, uint samplingRate, ushort bitPerSample, short[] data)
            {
                if(channels != 1)
                {
                    throw new ArgumentException("チャンネル数が1の音声データしか受け付けられません", "channels"); 
                }

                _samplingRate = samplingRate;

                foreach (var x0 in data)
                {
                    float x = Math.Abs(x0);
                    if (_mode)
                    {
                        ToSilentMode(x);
                    }
                    else
                    {
                        ToSoundMode(x); 
                    }
                    _count++;
                }
            }

            private void ToSilentMode(float value)
            {
                int keep_frame = (int)(_samplingRate * 1.0f); // 1.0秒音無しが続いたら無音区間に切り替える
                if (value < _silentThrethold)
                {
                    _keep++;
                    if (_keep > keep_frame)
                    {
                        _keep = 0;
                        _mode = false;
                        int end = _count;
                        _ranges.Add((_start, end));
                    }
                }
                else
                {
                    _keep = 0;
                }
            }

            private void ToSoundMode(float value)
            {
                if (value > _silentThrethold)
                {
                    int start_margin = (int)(_samplingRate * 0.5); // 開始時のマージンは0.5秒くらい
                    _mode = true;
                    _start = _count - start_margin;
                    if (_start < 0)
                    {
                        _start = 0;
                    }
                }

            }
        }
    }


}
