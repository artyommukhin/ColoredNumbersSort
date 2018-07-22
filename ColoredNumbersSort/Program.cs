using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.IO;
using Microsoft.Office.Interop.Word;
using Xceed.Words.NET;


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



            Color GetColorFromId(WdColorIndex colorIndex)
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
                    case WdColorIndex.wd:
                        return sortOrder[1];
                    case WdColorIndex.wdGreen:
                        return sortOrder[1];
                    case WdColorIndex.wdGreen:
                        return sortOrder[1];
                    case WdColorIndex.wdGreen:
                        return sortOrder[1];
                    case WdColorIndex.wdGreen:
                        return sortOrder[1];
                    case WdColorIndex.wdGreen:
                        return sortOrder[1];
                    case WdColorIndex.wdGreen:
                        return sortOrder[1];
                    default:
                        return sortOrder[14];

                }
            }

            
            
            List<int> orderedNumbers = new List<int>();

            List<ColoredNumber> coloredNumbers = new List<ColoredNumber>();

            Console.WriteLine("write filename");
            string filename = Console.ReadLine();

            DocX doc = DocX.Load(filename);

            foreach (var VARIABLE in doc.Text)
            {
                
            }

            
            Application application = new Application();
            Document document = application.Documents.Open(filename);

            foreach (Range character in document.Characters)
            {
                if (int.TryParse(character.ToString(),out int num))
                {
                    coloredNumbers.Add(new ColoredNumber(num, character.HighlightColorIndex));      
                }
            }



        }
    }
}
