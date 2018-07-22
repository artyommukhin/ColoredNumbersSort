using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.IO;
using Microsoft.Office.Interop.Word;
using Xceed.Words.NET;
/*using Spire.Doc;
using Spire.Doc.Documents;*/


namespace ColoredNumbersSort
{
    struct ColoredNumber
    {
        private int number;
        private Color color;

        public ColoredNumber(int number, Color color)
        {
            this.number = number;
            this.color = color;
        }
    }

    class Program
    {
        static void Main(string[] args)
        { 
            Dictionary<WdColorIndex,Color> colorsCompare = new Dictionary<WdColorIndex, Color>(15);
            colorsCompare.Add(WdColorIndex.wdDarkBlue, Color.DarkBlue);

            Color[] sortOrder =
            {
                Color.DarkBlue,
                Color.DarkGreen,
                Color.DarkCyan, 
                Color.DarkRed,
                Color.DarkMagenta,
                Color.Olive,
                Color.DarkGray,
                Color.Blue,
                Color.Green,
                Color.Cyan,
                Color.Red,
                Color.Magenta,
                Color.Yellow,
                Color.Gray,
                Color.White
            };
            
            Color GetColorById(WdColorIndex colorIndex)
            {
                switch (colorIndex)
                {
                    case WdColorIndex.wdDarkBlue:
                        return sortOrder[0];
                    case WdColorIndex.wdGreen:
                        return sortOrder[1];
                    case WdColorIndex.wdTeal:
                        return sortOrder[2];
                    case WdColorIndex.wdDarkRed:
                        return sortOrder[3];
                    case WdColorIndex.wdViolet:
                        return sortOrder[4];
                    case WdColorIndex.wdDarkYellow:
                        return sortOrder[5];
                    case WdColorIndex.wdGray50:
                        return sortOrder[6];
                    case WdColorIndex.wdBlue:
                        return sortOrder[7];
                    case WdColorIndex.wdBrightGreen:
                        return sortOrder[8];
                    case WdColorIndex.wdTurquoise:
                        return sortOrder[9];
                    case WdColorIndex.wdRed:
                        return sortOrder[10];
                    case WdColorIndex.wdPink:
                        return sortOrder[11];
                    case WdColorIndex.wdYellow:
                        return sortOrder[12];
                    case WdColorIndex.wdGray25:
                        return sortOrder[13];
                    case WdColorIndex.wdWhite:
                        return sortOrder[14];
                    default:
                        return sortOrder[14];

                }
            }

            
            
            List<int> orderedNumbers = new List<int>();

            List<ColoredNumber> coloredNumbers = new List<ColoredNumber>();

            Console.WriteLine("write filename");
            string filename = Console.ReadLine();
            
            
            //////Spire
           /* Document document = new Document();

            document.LoadFromFile(filename);

            TextSelection[] text = document*/









            /////////////////////////Xceed
            DocX doc = DocX.Load(filename);

            foreach (var VARIABLE in doc.Text)
            {
                
            }








            ///////////////////Interop
            Application application = new Application();
            Document document = application.Documents.Open(filename);

            string number = "";
            for (int i = 0; i < document.Characters.Count; i++)
            {
                Range currentChar = document.Characters[i];

                if (currentChar.HighlightColorIndex != WdColorIndex.wdNoHighlight)
                {
                    if (int.TryParse(currentChar.ToString(), out int num) &&
                       (number == "" || currentChar.HighlightColorIndex == document.Characters[i - 1].HighlightColorIndex))
                    {

                        coloredNumbers.Add(new ColoredNumber(num, GetColorById(currentChar.HighlightColorIndex)));
                    }
                }
            }


            void


        }
    }
}
