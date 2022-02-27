using System;
using System.Linq;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PPTXTools
{
    /// <summary>
    /// パワーポイントの情報を管理する
    /// </summary>
    public class PPTXLoader
    {
        /// <summary>
        /// パワーポイントのスライド情報
        /// </summary>
        public PPTXSlide[] PPTXSlides { get; private set; } 

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="path">パワーポイントのファイルのパス</param>
        public PPTXLoader(string path)
        {
            var ppt = new Application();
            var pres = ppt.Presentations;
            Presentation file = pres.Open(path, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
            PPTXSlides = new PPTXSlide[file.Slides.Count];
            float pre = 0.0f;
            for (int i = 0; i < file.Slides.Count; i++)
            {
                Slide slide = file.Slides[i + 1]; // スライドのIDは1から始まる（０からではない…）
                PPTXSlides[i] = new PPTXSlide(i+1, slide, pre);
                pre = PPTXSlides[i].EndTimeStamp;
            }
        }

        /// <summary>
        /// パワーポイントの情報を文字列化する
        /// </summary>
        /// <returns>文字列化されたパワーポイントの情報</returns>
        public override string ToString()
        {
            return string.Join("\n\n", PPTXSlides.Select(s => $"===== SLIDE ======\n{s}\n"));
        }
    }
}
