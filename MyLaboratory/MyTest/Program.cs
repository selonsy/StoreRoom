using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MsWord = Microsoft.Office.Interop.Word;

namespace MyTest
{
    class Program
    {
        static void Main(string[] args)
        {
            //try
            //{
            //    MsWord.Application wordApp = new MsWord.ApplicationClass();
            //    wordApp.Visible = true;
            //    MsWord.Document wordDoc = wordApp.Documents.Open("./智联招聘-test-word.doc");
            //    string ss= wordDoc.Paragraphs.Last.Range.Text;
            //}
            //catch (Exception ex)
            //{

            //}

            string a = "a";
            string b = "b";
            Swap( a, b);
            Console.WriteLine(a);
            Console.WriteLine(b);
            Console.ReadKey();
        }

        public static void Swap( string a,  string b)
        {
            string temp = a;
            a = b;
            b = temp;
        }
    }
}
