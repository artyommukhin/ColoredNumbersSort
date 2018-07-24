using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.IO;
using Microsoft.Office.Interop.Word;

namespace ColoredNumbersSort
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("write filename");
            string filename = Console.ReadLine();

            Application application = new Application();
            Document document = application.Documents.Open(filename);

            List<int> orderedNumbers = new List<int>();

            List<int> darkBlueNums = new List<int>();
            List<int> darkGreenNums = new List<int>();
            List<int> darkCyanNums = new List<int>();
            List<int> darkRedNums = new List<int>();
            List<int> darkMagentaNums = new List<int>();
            List<int> darkYellowNums = new List<int>();
            List<int> darkGrayNums = new List<int>();
            List<int> blueNums = new List<int>();
            List<int> greenNums = new List<int>();
            List<int> cyanNums = new List<int>();
            List<int> redNums = new List<int>();
            List<int> magentaNums = new List<int>();
            List<int> yellowNums = new List<int>();
            List<int> grayNums = new List<int>();
            List<int> whiteNums = new List<int>();

            List<int> otherNums = new List<int>();

            foreach (Range word in document.Words)
            {
                if (int.TryParse(word.Text, out int num))
                {
                    switch (word.HighlightColorIndex)
                    {
                        case WdColorIndex.wdDarkBlue:
                            darkBlueNums.Add(num);
                            break;
                        case WdColorIndex.wdGreen:
                            darkGreenNums.Add(num);
                            break;
                        case WdColorIndex.wdTeal:
                            darkCyanNums.Add(num);
                            break;
                        case WdColorIndex.wdDarkRed:
                            darkRedNums.Add(num);
                            break;
                        case WdColorIndex.wdViolet:
                            darkMagentaNums.Add(num);
                            break;
                        case WdColorIndex.wdDarkYellow:
                            darkYellowNums.Add(num);
                            break;
                        case WdColorIndex.wdGray50:
                            darkGrayNums.Add(num);
                            break;
                        case WdColorIndex.wdBlue:
                            blueNums.Add(num);
                            break;
                        case WdColorIndex.wdBrightGreen:
                            greenNums.Add(num);
                            break;
                        case WdColorIndex.wdTurquoise:
                            cyanNums.Add(num);
                            break;
                        case WdColorIndex.wdRed:
                            redNums.Add(num);
                            break;
                        case WdColorIndex.wdPink:
                            magentaNums.Add(num);
                            break;
                        case WdColorIndex.wdYellow:
                            yellowNums.Add(num);
                            break;
                        case WdColorIndex.wdGray25:
                            grayNums.Add(num);
                            break;
                        case WdColorIndex.wdNoHighlight:
                            whiteNums.Add(num);
                            break;
                        default:
                            otherNums.Add(num);
                            break;
                    }
                }
            }

            if (otherNums.Count > 0)
            {
                Console.WriteLine("Error: unexpected color found");
                Console.ReadKey();
                return;
            }

            document.Close();

            orderedNumbers.AddRange(darkBlueNums);
            orderedNumbers.AddRange(darkGreenNums);
            orderedNumbers.AddRange(darkCyanNums);
            orderedNumbers.AddRange(darkRedNums);
            orderedNumbers.AddRange(darkMagentaNums);
            orderedNumbers.AddRange(darkYellowNums);
            orderedNumbers.AddRange(darkGrayNums);
            orderedNumbers.AddRange(blueNums);
            orderedNumbers.AddRange(greenNums);
            orderedNumbers.AddRange(cyanNums);
            orderedNumbers.AddRange(redNums);
            orderedNumbers.AddRange(magentaNums);
            orderedNumbers.AddRange(yellowNums);
            orderedNumbers.AddRange(grayNums);
            orderedNumbers.AddRange(whiteNums);

            Console.Write("\nArranged numbers:\n"+string.Join(",", orderedNumbers));

            Console.ReadKey();
        }
    }
}
