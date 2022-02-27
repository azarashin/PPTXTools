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
    /// 軽量であるが、キャッシュを生成するための時間がかかる。
    /// 但し、クラスの構成は比較的シンプルでメンテナンスしやすい。
    /// </summary>
    public class WaveSplitter : IWaveSplitter
    {
        private string _path; 

        public (int, int)[] Ranges { get; private set; }
        public (float, float)[] RangesSec { get; private set; }
        public const float MinDurationSec = 1.0f;


        private float _silentThrethold; // 無音区間かどうかを区別するための閾値
        private int _keep = 0;
        private bool _mode = false; // true: 有音, falsle: 無音
        private int _count = 0;
        private int _start = -1;
        private List<(int, int)> _ranges;
        private uint _samplingRate;
        private Dictionary<float, (int, int)[]> _rangesMap;
        private Dictionary<float, (float, float)[]> _rangesSecMap;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="path">有音区間を抽出する対象の音声ファイルのパス</param>
        public WaveSplitter(string path)
        {
            _path = path;
            _rangesMap = new Dictionary<float, (int, int)[]>();
            _rangesSecMap = new Dictionary<float, (float, float)[]>(); 
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
            WaveData data = new WaveData(_path);
            WaveAttribute attr = new WaveAttribute(data);
            float silentWeight;
            float speed = 0.05f;
            float bias = 0.001f;
            int count = 0;
            int maxCount = 10;
            int targetCount; 

            do
            {
                silentWeight = speed * (maxCount - count) + bias;
                attr.UpdateSilentParameter(silentWeight);
                if(!_rangesMap.ContainsKey(silentWeight))
                {
                    _rangesMap[silentWeight] = Split(data, attr, _samplingRate);
                    _rangesSecMap[silentWeight] = _rangesMap[silentWeight].Select(s => (s.Item1 / (float)_samplingRate, s.Item2 / (float)_samplingRate)).ToArray();
                }
                Ranges = _rangesMap[silentWeight];
                RangesSec = _rangesSecMap[silentWeight];
                count++;
                targetCount = RangesSec
                    .Where(s => s.Item1 >= start && s.Item2 <= end)
                    .Count();
            } while (targetCount < expectedLength && count <= maxCount);

            return Ranges.Length; 
        }


        private void ScanSplitParameter(ushort channels, uint samplingRate, ushort bitPerSample, short[] data)
        {
            _samplingRate = samplingRate;
            int keepFrame = (int)(_samplingRate * 1.0f); // 1.0秒音無しが続いたら無音区間に切り替える
            int startMargin = (int)(_samplingRate * 0.5); // 開始時のマージンは0.5秒くらい

            foreach (var x0 in data)
            {
                float x = Math.Abs(x0);
                if (_mode)
                {
                    if (x < _silentThrethold)
                    {
                        _keep++;
                        if (_keep > keepFrame)
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
                else
                {
                    if (x > _silentThrethold)
                    {
                        _mode = true;
                        _start = _count - startMargin;
                        if (_start < 0)
                        {
                            _start = 0;
                        }
                    }

                }
                _count++;

            }
        }

        private (int, int)[] Split(WaveData data, WaveAttribute attr, uint samplingRate)
        {
            _keep = 0;
            _mode = false;
            _count = 0;
            _start = -1;

            _silentThrethold = attr.SilentParameter;

            _ranges = new List<(int, int)>();

            data.Scan(ScanSplitParameter);

            _ranges = _ranges.Where(s => s.Item2 - s.Item1 > MinDurationSec * samplingRate).ToList(); 


            if (_ranges.Count() == 0)
            {
                return new (int, int)[0]; 
            }
            List<(int, int)> range_final = new List<(int, int)>();
            range_final.Add(_ranges[0]);
            int threshold_sec = 1; // この秒数未満の区間は一つにまとめる
            int threshold = threshold_sec * (int)_samplingRate; 
            for(int i=1;i<_ranges.Count();i++)
            {
                (int, int) lst = range_final.Last();
                (int, int) cur = _ranges[i];
                if (cur.Item2 - lst.Item1 < threshold) {
                    range_final.RemoveAt(range_final.Count() - 1);
                    range_final.Add((lst.Item1, cur.Item2));
                } else
                {
                    range_final.Add(cur); 
                }
            }
            return range_final.ToArray(); 
        }

        /// <summary>
        /// 抽出された有音区間を取得する
        /// </summary>
        /// <returns>(有音区間の開始位置（秒）, 有音区間の終端位置(秒))の配列</returns>
        public (float, float)[] GetRangesSec()
        {
            return RangesSec;
        }
    }
}
