using MeCab;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace PPTXTools
{
    /// <summary>
    /// パワーポイントのノートデータと音声ファイルとから字幕データをSRTファイルとして生成する
    /// </summary>
    public class SRTGenerator
    {
        /// <summary>
        /// (区間開始位置(秒),区間終端位置(秒), 字幕)
        /// </summary>
        public ((float, float), string)[] NoteToTimestamp { get; private set; }

        /// <summary>
        /// 改行までの最大文字数
        /// </summary>
        public const int WrapLength = 40; 

        private MeCabTagger _tagger;
        private TextWrapper _wrapper; 
        private int[] _lengthOfNumber;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="pptx">パワーポイントのファイルのパス</param>
        /// <param name="wave">音声ファイルのパス</param>
        public SRTGenerator(PPTXLoader pptx, IWaveSplitter wave)
        {
            _wrapper = new TextWrapper();
            _tagger = MeCabTagger.Create();
            _lengthOfNumber = new int[] { 2, 2, 1, 2, 2, 1, 2, 2, 2, 2 };

            List<((float, float), string)> duration_notes = new List<((float, float), string)>();
            foreach (PPTXSlide info in pptx.PPTXSlides)
            {
                ((float, float), string)[] sub = GetNoteToTimestamp(info, wave);
                duration_notes.AddRange(sub);
            }
            NoteToTimestamp = duration_notes.OrderBy(s => s.Item1.Item1).ToArray();
        }

        /// <summary>
        /// SRTファイルに字幕データを出力する
        /// </summary>
        /// <param name="path">出力先のパス</param>
        public void DumpSRT(string path)
        {
            using (StreamWriter sw = new StreamWriter(path))
            {
                int id = 0;
                foreach (((float, float) duration, string note) in NoteToTimestamp)
                {
                    sw.WriteLine($"{id}");
                    sw.WriteLine($"{SecToSRTTime(duration.Item1)} --> {SecToSRTTime(duration.Item2)}");
                    sw.WriteLine($"{_wrapper.Wrap(note)}\n");
                    id++;
                }
            }
        }


        private int LengthOfWordPronounsation(string word)
        {
            float fvalue;
            int ivalue; 
            if(!float.TryParse(word, out fvalue))
            {
                return word.Replace("ャ", "").Replace("ュ", "").Replace("ョ", "").Length;
            }
            int length = 0; 
            if(word[0] == '-')
            {
                length = 4; // "マイナス"
                word = word.Substring(1); 
            }
            if(int.TryParse(word, out ivalue))
            {
                length += word.Sum(s => _lengthOfNumber[s - '0']); // 各桁の数字の読みの長さを合計する
                length += (word.Length - 1) * 2; // 千・百・十とか万・億・兆の部分の読みの長さを合計する
                return length; 
            }
            // 小数
            string[] floatNumber = word.Split('.');
            length += floatNumber[0].Sum(s => _lengthOfNumber[s - '0']); // 各桁の数字の読みの長さを合計する
            length += (word.Length - 1) * 2; // 千・百・十とか万・億・兆の部分の読みの長さを合計する
            length += 2; // 小数点の「てん」の２文字
            length += floatNumber[1].Sum(s => _lengthOfNumber[s - '0']); // 小数部分の各桁の数字の読みの長さを合計する
            return length; 


        }

        private int LengthOfSequencePronounsation(string text)
        {
            int sum = 0;
            int length; 
            foreach(var t in _tagger.ParseToNodes(text))
            {
                string[] attribs = t.Feature.Split(',');
                if (attribs[0] == "BOS/EOS")
                {
                    continue; 
                }
                if (attribs.Length < 8)
                {
                    length = LengthOfWordPronounsation(t.Surface); 
                } else
                {
                    length = LengthOfWordPronounsation(attribs[7]);
                }
                sum += length;
                // Console.WriteLine($"{t.Surface} - {length}"); 
            }
            return sum; 
        }

        private string SecToSRTTime(float tm)
        {
            int h = (int)(tm / 3600);
            int m = (int)(tm / 60) % 60;
            int s = (int)(tm) % 60;
            float msec = tm - (int)tm;
            string smsec = $"{msec:F3}".Substring(2);
            return $"{h:D2}:{m:D2}:{s:D2},{smsec}";
        }

        private string Wrap(string source)
        {
            string ret = ""; 
            while(source != "")
            {
                int length = WrapLength; 
                if(source.Length < length)
                {
                    length = source.Length; 
                }
                
                ret += source.Substring(0, length);
                source = source.Substring(length); 
                if(source != "")
                {
                    ret += "\n"; 
                }
            }
            return ret; 
        }

        private (float, float) NearestTimestampWithRate(((float, float), float)[] rateMap, float rate)
        {
            (float, float) ret = (0.0f, 0.0f);
            float min = float.MaxValue; 
            foreach (((float, float) duration, float drate) in rateMap)
            {
                float diff = Math.Abs(drate - rate); 
                if(min > diff)
                {
                    min = diff;
                    ret = duration; 
                }
            }
            return ret; 
        }

        private (float, float) NextDuration((float, float)[] all, (float, float) current)
        {
            foreach(var t in all)
            {
                if(t.Item1 > current.Item1)
                {
                    return t;
                } 
            }
            return all.Last();
        }

        private (float, float) PrevDuration((float, float)[] all, (float, float) current)
        {
            foreach (var t in all.Reverse())
            {
                if (t.Item1 < current.Item1)
                {
                    return t;
                }
            }
            return all.First();
        }

        private ((float, float), string)[] GetNoteToTimestampWithoutWave(PPTXSlide info)
        {
            string note = Regex.Replace(info.NoteText.Replace("\r", "\n"), "\n\n+", "\n\n");
            string[] notes = note
                .Split(new string[] { "\n\n" }, StringSplitOptions.RemoveEmptyEntries)
                .Select(s => s.Replace("\n", ""))
                .ToArray();

            List<(int, string)> noteAt = new List<(int, string)>();
            noteAt.Add((0, notes[0]));
            int tm = 0;
            for (int i = 1; i < notes.Length; i++)
            {
                tm += LengthOfSequencePronounsation(notes[i - 1]);
                noteAt.Add((tm, notes[i]));
            }
            noteAt.Add((tm + LengthOfSequencePronounsation(notes.Last()), "")); // 番兵
            int totalLength = notes.Sum(s => LengthOfSequencePronounsation(s));
            (string, float)[] rateMapNote = noteAt
                .Select(s => (s.Item2, (float)(s.Item1) / (float)(totalLength))).ToArray();

            float start = info.TimeStamp;
            float end = info.EndTimeStamp;
            return Enumerable.Range(0, notes.Length)
                .Select(s => ((
                    (start + (end - start) * rateMapNote[s].Item2,
                    start + (end - start) * rateMapNote[s + 1].Item2))
                    , notes[s])).ToArray();
        }

        private ((float, float), string)[] GetNoteToTimestamp(PPTXSlide info, IWaveSplitter wave)
        {
            if(info.Hidden)
            { // 非表示スライドの場合
                return new ((float, float), string)[0];
            }
            if (info.TimeStamp == info.EndTimeStamp)
            {
                // 記録のないスライドが含まれている場合、パワポと動画でタイムスタンプがずれる
                // 問題があるため（パワポ側の問題なので問題回避不可）、
                // エラーメッセージを出す(非表示スライドは対象外)。
                throw new NoRecordSlideException(info.PageNumber);
            }

            string note = Regex.Replace(info.NoteText.Replace("\r", "\n"), "\n\n+", "\n\n");
            string[] notes = note.Split(new string[] { "\n\n" }, StringSplitOptions.RemoveEmptyEntries);

            if(notes.Length == 0)
            {
                return new ((float, float), string)[0]; 
            }

            float margin = 1.2f; 

            wave.Scan((int)(notes.Length * margin), info.TimeStamp, info.EndTimeStamp); // セリフの分割数以上に音声データが分割されるように指定する

            if(wave.GetRangesSec() == null || wave.GetRangesSec().Length < notes.Length)
            {
                return GetNoteToTimestampWithoutWave(info); 
            }

            (float, float)[] target = wave.GetRangesSec()
                .Where(s => s.Item1 >= info.TimeStamp && s.Item1 <= info.EndTimeStamp)
                .ToArray();
            ((float, float), float)[] rateMapWave = target
                .Select(s => ((s.Item1, s.Item2), (s.Item1 - target.First().Item1) / (target.Last().Item2 - target.First().Item1))).ToArray();

            List<(int, string)> noteAt = new List<(int, string)>();
            noteAt.Add((0, notes[0]));
            int tm = 0;
            for (int i = 1; i < notes.Length; i++)
            {
                tm += LengthOfSequencePronounsation(notes[i - 1]);
                noteAt.Add((tm, notes[i]));
            }
            noteAt.Add((tm + LengthOfSequencePronounsation(notes.Last()), "")); // 番兵
            int totalLength = notes.Sum(s => LengthOfSequencePronounsation(s));
            (string, float)[] rateMapNote = noteAt
                .Select(s => (s.Item2, (float)(s.Item1) / (float)(totalLength))).ToArray();

            if (target.Length == 0)
            {
                return GetNoteToTimestamp(info, notes, rateMapNote);
            }
            if (notes.Length > target.Length)
            {
                return GetNoteToTimestamp(target, notes);
            }

            // 音声検出による話し始めと話し終わりの位置になるように、単語列による相対位置情報を補正する
            float wavStart = target.First().Item1;
            float wavEnd = target.Last().Item2;
            float wordStart = info.TimeStamp;
            float wordEnd = info.EndTimeStamp;
            float adjustBias = (wavStart - wordStart) / (wordEnd - wordStart);
            float adjustScale = (wavEnd - wordStart) / (wordEnd - wordStart);


            ((float, float), string)[] noteToTimestamp = rateMapNote
                .Take(rateMapNote.Length - 1) // 番兵は除去する
                .Select(s => (NearestTimestampWithRate(rateMapWave, Adjust(s.Item2, adjustScale, adjustBias)), s.Item1)).ToArray();
            for (int i = 1; i < noteToTimestamp.Length; i++)
            {
                if (noteToTimestamp[i].Item1.Item1 <= noteToTimestamp[i - 1].Item1.Item1)
                {
                    // 前に詰まっていたら少しずつ後ろにずらす
                    noteToTimestamp[i].Item1 = NextDuration(target, noteToTimestamp[i - 1].Item1);
                }
            }

            for (int i = noteToTimestamp.Length - 1; i > 0; i--)
            {
                if (noteToTimestamp[i].Item1.Item1 <= noteToTimestamp[i - 1].Item1.Item1)
                {
                    // 後ろに詰まっていたら少しずつ前にずらす
                    noteToTimestamp[i - 1].Item1 = PrevDuration(target, noteToTimestamp[i].Item1);
                }
            }

            for (int i = 0; i < noteToTimestamp.Length - 1; i++)
            {
                noteToTimestamp[i].Item1.Item2 = noteToTimestamp[i + 1].Item1.Item1; 
            }
            noteToTimestamp[noteToTimestamp.Length - 1].Item1.Item2 = info.EndTimeStamp; 
            return noteToTimestamp; 
        }

        private float Adjust(float source, float scale, float bias)
        {
            return source * scale + bias; 
        }

        private ((float, float), string)[] GetNoteToTimestamp((float, float)[] target, string[] notes)
        {
            float start = target.First().Item1;
            float end = target.Last().Item2;
            return Enumerable.Range(0, notes.Length)
                .Select(s => ((
                    (start + (end - start) * s / notes.Length,
                    start + (end - start) * (s + 1) / notes.Length))
                    , notes[s])).ToArray();
        }

        private ((float, float), string)[] GetNoteToTimestamp(PPTXSlide info, string[] notes, (string, float)[] rateMapNote)
        {
            float start = info.TimeStamp;
            float end = info.EndTimeStamp;
            return Enumerable.Range(0, notes.Length)
                .Select(s => ((
                    (start + (end - start) * rateMapNote[s].Item2,
                    start + (end - start) * rateMapNote[s + 1].Item2))
                    , notes[s])).ToArray();
        }
    }
}
