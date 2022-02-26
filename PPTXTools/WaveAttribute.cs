using PPTXTools.Wave;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PPTXTools
{
    /// <summary>
    /// 音声データを解析して得られた属性情報を管理する
    /// </summary>
    public class WaveAttribute
    {
        /// <summary>
        /// 無音区間かどうかを区別するための閾値
        /// </summary>
        public float SilentParameter { get; private set; }

        private float _cur = 0.0f;
        private int _cnt = 0;
        private float _max = 0.0f;
        private float _min = 32767.0f;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="data">音声データを読み込むためのWaveData インスタンス</param>
        public WaveAttribute(WaveData data)
        {
            float silentWeight = 0.002f;

            data.Scan(ScanSilentParameter); 
            SilentParameter = (int)(_min + (_max - _min) * silentWeight);
        }

        /// <summary>
        /// 重みパラメータを指定して、無音区間かどうかを区別するための閾値を更新する
        /// </summary>
        /// <param name="silentWeight">新しい重みパラメータ。0.0f～1.0fの範囲で指定する。
        /// 0.0f 未満の場合は0.0f に、1.0f を超える場合は1.0f に修正される。</param>
        /// <returns>更新された閾値</returns>
        public float UpdateSilentParameter(float silentWeight)
        {
            silentWeight = Math.Min(Math.Max(0.0f, silentWeight), 1.0f);
            // silentWeight が大きいほど無音と判定するようになるので、細かく分割されやすくなる
            SilentParameter = (int)(_min + (_max - _min) * silentWeight);
            return SilentParameter; 
        }


        private void ScanSilentParameter(ushort channels, uint samplingRate, ushort bitPerSample, short[] data)
        {
            float smooth = 0.2f;
            int keep = 100;
            foreach (short x0 in data)
            {

                short x = Math.Abs(x0);
                _cnt++;
                _cur += (x - _cur) * smooth;
                if (_cnt > keep)
                {
                    if (_min > _cur)
                    {
                        _min = _cur;
                    }
                    if (_max < _cur)
                    {
                        _max = _cur;
                    }
                }

            }
        }
    }
}
