using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PPTXTools
{
    class PPTXErrorChecker
    {
        string _message = ""; 

        public PPTXErrorChecker(Presentation data)
        {
            _message = ""; 
            _message += MayBeInvalidAdvanceTime(data);
            _message += ContainsMediaAnimation(data); 
        }

        public override string ToString()
        {
            return _message; 
        }

        private string ContainsMediaAnimation(Presentation data)
        {
            string ret = "";
            Dictionary<int, int> count = new Dictionary<int, int>(); 
            for (int i = 0; i < data.Slides.Count; i++)
            {
                Slide slide = data.Slides[i + 1];
                count[i + 1] = 0;
                for(int j = 0;j < slide.TimeLine.MainSequence.Count;j++)
                {
                    MsoAnimEffect type = slide.TimeLine.MainSequence[j + 1].EffectType; 
                    float animDuration = slide.TimeLine.MainSequence[j + 1].Timing.Duration; 
                    if (type == MsoAnimEffect.msoAnimEffectMediaPause
                        || type == MsoAnimEffect.msoAnimEffectMediaPlay
                        || type == MsoAnimEffect.msoAnimEffectMediaPlayFromBookmark
                        || type == MsoAnimEffect.msoAnimEffectMediaStop)
                    {
                        float transDuration = slide.SlideShowTransition.AdvanceTime + slide.SlideShowTransition.Duration;
                        if (transDuration < animDuration)
                        {
                            count[i + 1]++;
                        }
                    }
                }
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
            return ret; 
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
