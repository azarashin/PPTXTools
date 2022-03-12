using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PPTXTools.Wave;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PPTXTools
{
    public class PPTXErrorChecker
    {
        string _message = "";
        float _totalWaveLength; 

        public float TotalWaveLength()
        {
            return _totalWaveLength; 
        }

        public PPTXErrorChecker(string path, string wavePath)
        {
            _totalWaveLength = 0.0f;
            WaveData wd = new WaveData(wavePath);
            wd.Scan(WaveScan);

            var ppt = new Application();
            var pres = ppt.Presentations;
            Presentation data = pres.Open(path, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);

            _message = ""; 
            _message += MayBeInvalidAdvanceTime(data);
            _message += ContainsMediaAnimation(data);



        }

        private void WaveScan(ushort channels, uint samplingRate, ushort bitPerSample, short[] data)
        {
            float dur = (data.Length / channels) / (float)samplingRate;
            _totalWaveLength += dur; 
        }

        public override string ToString()
        {
            return _message; 
        }

        private string ContainsMediaAnimation(Presentation data)
        {
            string ret = "";
            float totalDuration = 0.0f; 
            Dictionary<int, int> count = new Dictionary<int, int>();
            Dictionary<string, float> timeline = new Dictionary<string, float>(); 
            timeline["1"] = 0.0f; 
            for (int i = 0; i < data.Slides.Count; i++)
            {
                Slide slide = data.Slides[i + 1];
                count[i + 1] = 0;
                float transDuration = slide.SlideShowTransition.AdvanceTime + slide.SlideShowTransition.Duration;
                for (int j = 0;j < slide.TimeLine.MainSequence.Count;j++)
                {
                    MsoAnimEffect type = slide.TimeLine.MainSequence[j + 1].EffectType; 
                    float animDuration = slide.TimeLine.MainSequence[j + 1].Timing.Duration; 
                    if (type == MsoAnimEffect.msoAnimEffectMediaPause
                        || type == MsoAnimEffect.msoAnimEffectMediaPlay
                        || type == MsoAnimEffect.msoAnimEffectMediaPlayFromBookmark
                        || type == MsoAnimEffect.msoAnimEffectMediaStop)
                    {
                        if (transDuration < animDuration)
                        {
                            count[i + 2]++;
                        }
                    }
                }
                float next = timeline.Last().Value + transDuration;
                if(i < data.Slides.Count - 1)
                {
                    timeline[$"{i + 2}"] = next;
                }
                else
                {
                    timeline[$"Last"] = next;
                }
                totalDuration += transDuration;
            }
            count = count
                .Where(s => s.Value > 0)
                .ToDictionary(s => s.Key, s => s.Value);
            if(count.Count() > 0)
            {
                ret += "\n★警告：以下のスライドにはメディアの再生・停止などメディア関連のアニメーション制御が含まれており、\n" +
                    "このメディアの再生時間がスライドの画面切り替えのタイミングより長くなっています。\n" +
                    "これらのアニメーション制御により、動画生成時に不要な画像フレームが生成され、字幕のタイミングをずらす恐れがあります。\n";
                ret += string.Join(",", count.Select(s => $"スライド番号[{s.Key}]:{s.Value}箇所\n"));
            }

            float thresholdDuration = 0.1f; 
            if(Math.Abs(_totalWaveLength - totalDuration) / data.Slides.Count > thresholdDuration)
            {
                ret += $"\n★警告：パワポの合計再生時間と動画の再生時間とが乖離しています。\n" +
                    $"メディアの再生終了タイミングが画面切り替えのタイミングよりも後になっている可能性があります。\n" +
                    $"パワポのスライドショーを実行し、スライドの再生が終わっているのにページが切り替わっていないケースがないかどうかを確認してください。\n" +
                    $"スライドの枚数: {data.Slides.Count}\n" +
                    $"パワポの合計再生時間：{totalDuration}\n" +
                    $"動画の再生時間：{_totalWaveLength}\n" +
                    $"\n";
                ret += string.Join(",", timeline.Select(s => $"スライド番号[{s.Key}]:{TimeString(s.Value)}\n"));
            }

            return ret; 
        }

        private string TimeString(float sec)
        {
            int hours = (int)(sec / 3600);
            int minutes = (int)(sec / 60) - hours * 60;
            float seconds = sec - hours * 3600 - minutes * 60;
            return $"{hours}:{minutes}:{seconds}";

        }

        private string MayBeInvalidAdvanceTime(Presentation data)
        {
            string ret = "";
            List<int> noAdvanceOnTime = new List<int>(); 
            Dictionary<float, List<int>> count = new Dictionary<float, List<int>>(); 
            for (int i = 0; i < data.Slides.Count; i++)
            {
                Slide slide = data.Slides[i + 1];
                if(slide.SlideShowTransition.AdvanceOnTime == Microsoft.Office.Core.MsoTriState.msoFalse)
                {
                    noAdvanceOnTime.Add(i + 1);
                }
                float dur = slide.SlideShowTransition.AdvanceTime; 
                if(!count.ContainsKey(dur))
                {
                    count[dur] = new List<int>(); 
                }
                count[dur].Add(i + 1); 
            }
            if(noAdvanceOnTime.Count() > 0)
            {
                ret += $"\n☆警告：以下のページは「画面切り替えのタイミング」が自動設定になっていません。これは意図通りのものですか？\n";
                ret += string.Join(",", noAdvanceOnTime.Select(s => s.ToString()));
                ret += "\n";
            }
            float[] listOfSameAdvanceOnTime = count
                .Where(s => s.Value.Count() >= 2)
                .Select(s => s.Key)
                .ToArray(); 
            if(listOfSameAdvanceOnTime.Count() > 0)
            {
                ret += $"\n☆警告：複数のページの画面切り替えのタイミングが同じ値になっています。タイミングの一括設定が行われ、本来のタイミング値から変更されている可能性があります。\n";
                foreach(var tm in listOfSameAdvanceOnTime)
                {
                    string tmList = string.Join(",", count[tm].Select(s => s.ToString())); 
                    ret += $"画面切り替えのタイミング：{tm}, 対象スライド番号：[{tmList}]\n";
                    
                }
            }
            return ret; 
        }
    }
}
