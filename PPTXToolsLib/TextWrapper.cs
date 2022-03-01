using MeCab;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PPTXTools
{
    /// <summary>
    /// テキストの折り返し処理を実行する。日本語が対象。
    /// </summary>
    public class TextWrapper
    {
        public const int MaxLength = 40; // 字幕の一行当たりの最大文字数
        public const int MinLength = 32; // 字幕の一行当たりの最小文字数
        private MeCabTagger _tagger;
        private StreamWriter _swDebug;
        private string[] postCharactors = new string[] { "、", "。", "」", "】", "』", ")", "）", ">", "＞", "》", "≫", "〕", "］", "｝" };

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public TextWrapper(string noLineBreakBefore)
        {
            _tagger = MeCabTagger.Create();
            _swDebug = null; // new StreamWriter("log.txt");
            if(noLineBreakBefore != null && noLineBreakBefore != "")
            {
                postCharactors = Enumerable.Range(0, noLineBreakBefore.Length).Select(s => noLineBreakBefore.Substring(s, 1)).ToArray();
            }
        }

        /// <summary>
        /// テキストの折り返し処理を実行する
        /// </summary>
        /// <param name="source">折り返し対象のテキスト</param>
        /// <returns>折り返されたテキスト</returns>
        public string Wrap(string source)
        {
            MeCabNode[] words = _tagger.ParseToNodes(source)
                .Skip(1) // 先頭は文全体なので除外する
                .ToArray();
            words = words.Take(words.Length - 1).ToArray(); // 末尾はEOS/BOS なので除外。

            List<List<MeCabNode>> nodeLists = new List<List<MeCabNode>>();
            nodeLists.Add(words.ToList());

            for (int i = 0; i < words.Length + 1; i++)
            {
                List<List<MeCabNode>> next = new List<List<MeCabNode>>();

                foreach (var nodes in nodeLists)
                {
                    if (i >= nodes.Count())
                    {
                        break;
                    }
                    MeCabNode node = nodes[i];
                    if (node == null)
                    {
                        continue;
                    }
                    int length = SurfaceLength(nodes, i);
                    if (length <= MaxLength && length > MinLength)
                    {
                        List<MeCabNode> nw = nodes.ToList();
                        nw.Insert(i + 1, null); // 改行を追加したノードリストを新たに生成する
                        next.Add(nw);
                    }
                    else if (length > MaxLength)
                    {
                        nodes.Insert(i + 1, null); // このノードリストに改行を追加する
                    }
                }
                nodeLists.AddRange(next);
            }

            if (_swDebug != null)
            {
                foreach (var lst in nodeLists)
                {
                    _swDebug.WriteLine($"Score: {Score(lst)}");
                    _swDebug.WriteLine($"{string.Join("", lst.Select(s => NodeToSurface(s)))}");
                }
            }
            List<MeCabNode> best = nodeLists.OrderBy(s => Score(s)).Last();
            string ret = string.Join("", best.Select(s => NodeToSurface(s)));
            return ret;
        }

        private int SurfaceLength(List<MeCabNode> words, int index)
        {
            int start = 0; // 先頭
            for (int i=words.Count() - 1;i>=0;i--) // 改行位置を探す
            {
                if(words[i] == null)
                {
                    start = i + 1; // null の次
                    break; 
                }
            }
            int ret = Enumerable.Range(start, index - start + 1).Sum(s => words[s].Length);
            return ret; 
        }

        private string NodeToSurface(MeCabNode node)
        {
            if(node == null)
            {
                return "\n";
            } else
            {
                return node.Surface; 
            }
        }

        private float Score(List<MeCabNode> nodes)
        {
            List<int> wrapPos = new List<int>();
            List<int> lengthes = new List<int>();
            for(int i=0;i<nodes.Count();i++)
            {
                if(nodes[i] == null)
                {
                    wrapPos.Add(i); 
                    if(wrapPos.Count() == 1)
                    {
                        lengthes.Add(Enumerable.Range(0, i).Sum(s => nodes[s].Length));
                    } else
                    {
                        lengthes.Add(Enumerable.Range(wrapPos.Last() + 1, i - wrapPos.Last()).Sum(s => nodes[s].Length));
                    }
                }
            }
            if(lengthes.Count() == 0)
            {
                lengthes.Add(nodes.Sum(s => s.Length)); 
            } else
            {
                lengthes.Add(Enumerable.Range(wrapPos.Last() + 1, nodes.Count() - 1 - wrapPos.Last()).Sum(s => nodes[s].Length));
            }

            float score = -wrapPos.Count();
            float bigPenalty = 255.0f;
            float varWeight = -0.1f;
            float auxiliaryScore = 5.0f; 
            float commaScore = 10.0f;
            float periodScore = 20.0f;
            float auxiliaryPenalty = 5.0f; // 後半が助詞・助動詞になるところで改行されたときのペナルティ
            float separatedNounsPenalty = 2.0f; // 名詞が続くところで改行されたときのペナルティ

            if(lengthes.Any(s => (s > MaxLength)))
            {
                return -bigPenalty; // 規定文字数を超えていたらアウト
            }

            foreach (int index in wrapPos)
            {
                if(index == 0 || index == nodes.Count() - 1)
                {
                    score -= bigPenalty; 
                    continue; 
                }
                if(nodes[index-1].Surface == "、")
                {
                    score += commaScore; 
                }
                if (nodes[index - 1].Surface == "。")
                {
                    score += periodScore;
                }
                if (postCharactors.Contains(nodes[index + 1].Surface)) // 改行の後に句読点などがきてはいけない
                {
                    score -= bigPenalty;
                }
                string preType = nodes[index - 1].Feature.Split(',')[0];
                string postType = nodes[index + 1].Feature.Split(',')[0];
                if(postType == "助詞" || postType == "助動詞")
                {
                    score -= auxiliaryPenalty; 
                } else if (preType == "助詞" || preType == "助動詞")
                {
                    score += auxiliaryScore;
                }
                if ((preType == "名詞") && (postType == "名詞"))
                {
                    score -= separatedNounsPenalty;
                }
            }

            float lengthAverage = lengthes.Sum(s => s) / lengthes.Count();
            float lengthAverage2 = lengthes.Sum(s => s * s) / lengthes.Count();
            float vr = lengthAverage2 - lengthAverage * lengthAverage;

            score += varWeight * vr / MaxLength;

            return score; 
        }

    }
}
