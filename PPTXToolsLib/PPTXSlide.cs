using System;
using System.Linq;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PPTXTools
{
    /// <summary>
    /// パワーポイントのスライドの情報を管理する
    /// </summary>
    public class PPTXSlide
    {
        /// <summary>
        /// スライドの開始タイミング(秒)
        /// </summary>
        public float TimeStamp { get; private set; }

        /// <summary>
        /// スライドの終了タイミング(秒)
        /// </summary>
        public float EndTimeStamp { get; private set; }

        /// <summary>
        /// スライド中にある各テキスト領域のテキスト
        /// </summary>
        public string[] TextContents { get; private set; }

        /// <summary>
        /// スライド中のコメント群
        /// </summary>
        public string[] Comments { get; private set; }

        /// <summary>
        /// スライドに付与されたノートテキスト
        /// </summary>
        public string NoteText { get; private set; }

        /// <summary>
        /// スライド番号（先頭の番号は1）
        /// </summary>
        public int PageNumber { get; private set; }

        /// <summary>
        /// スライド非表示かどうか
        /// </summary>
        public bool Hidden { get; private set; }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="slide">スライド情報</param>
        /// <param name="timeStampOrigin">スライドの開始タイミング(秒)</param>
        public PPTXSlide(int pageNumber, Slide slide, float timeStampOrigin)
        {
            TimeStamp = timeStampOrigin;
            PageNumber = pageNumber;
            Hidden = (slide.SlideShowTransition.Hidden == MsoTriState.msoTrue);
            if (!Hidden)
            {
                EndTimeStamp = timeStampOrigin + slide.SlideShowTransition.AdvanceTime + slide.SlideShowTransition.Duration;
            } else
            {
                EndTimeStamp = timeStampOrigin;
            }

            TextContents = Enumerable.Range(0, slide.Shapes.Count)
                .Where(s => slide.Shapes[s + 1].HasTextFrame != 0)
                .Select(s => slide.Shapes[s + 1].TextFrame.TextRange.Text.Replace("\r", "\n"))
                .ToArray();


            Comments = Enumerable.Range(0, slide.Comments.Count)
                .Select(s => slide.Comments[s + 1].Text)
                .ToArray();

            NoteText = "";
            if (slide.HasNotesPage != 0 && slide.NotesPage.Count == 1)
            {
                Slide note = slide.NotesPage[1];
                if (note.Shapes.Placeholders.Count == 2)
                {
                    NoteText = note.Shapes.Placeholders[2].TextFrame.TextRange.Text;
                }
            }
        }

        /// <summary>
        /// スライド情報を文字列化する
        /// </summary>
        /// <returns>文字列化されたスライド情報</returns>
        public override string ToString()
        {
            string ret = "";
            ret += $"TimeStamp: {TimeStamp} - {EndTimeStamp}\n";
            ret += string.Join("\n", TextContents.Select(s => $"--- TEXT ---\n{s}\n"));
            ret += "\n";
            ret += string.Join("\n", Comments.Select(s => $"--- COMMENT ---\n{s}\n"));
            ret += "\n";
            ret += $"--- COMMENT ---\n{NoteText}\n";
            return ret; 
        }

        public void Adjust(float pptxLength, float waveLength)
        {
            TimeStamp *= waveLength / pptxLength;
            EndTimeStamp *= waveLength / pptxLength;
        }

        private float SumOfTimelineDuration(Slide slide)
        {
            int countMain = slide.TimeLine.MainSequence.Count;
            float sum = 0.0f; 
            for(int i=0;i<countMain;i++)
            {
                sum += slide.TimeLine.MainSequence[i + 1].Timing.Duration;
            }
            return sum; 
        }

    }
}
