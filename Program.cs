using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelObj
{
    class Excel
    {
        public ulong Kolumny, I;
        //public int I;
        public ulong[] Tab;

        public void Zmiana_Danych(ulong kolumny, ulong i)
        {
            Kolumny = kolumny;
            I = i;
        }

        public void Zmiana_Tablicy(ulong a)
        {
            Tab[I] = a;
            I++;
        }
        public ulong Czytanie_Tablicy(int i)
        {
            return Tab[i];

        }
               
        public void PobranieDanych(ulong kolumny)
        {
            Console.WriteLine("Prosze podac o ktora kolumne chodzi: \n");
            kolumny = Convert.ToUInt64(Console.ReadLine());
            if (kolumny > 18446744073709551615)
                {
                    Console.WriteLine("Przepraszam najwieksza liczba jest 18 446 744 073 709 551 615");
                    Console.ReadLine();
                    Zmiana_Danych(18446744073709551615, 0);
                }
                else
                   Zmiana_Danych(kolumny, 0);
        }
        
        public void Obliczanie(ulong kolumny)
        {
            while (kolumny > 0)
            {
                Zmiana_Tablicy(kolumny % 26);
                kolumny = kolumny / 26;
            }
        }
        
        public void Wypisanie(ulong kolumny, ulong i)
        {
             Console.WriteLine("Szukasz kolumny ");
             int a = Convert.ToInt32(i);
            for (int j = (a - 1); j >= 0; j--)
            {
                if (Czytanie_Tablicy(j) > 0)
                    Console.Write((char)(Czytanie_Tablicy(j) + 64));
                else
                    Console.Write(Czytanie_Tablicy(j));
            }
            //Console.WriteLine("\n");
            Console.ReadLine();

            
        }
        
       

        static void Main(string[] args)
        {
            Excel Ex = new Excel();
            Ex.Kolumny = 0;
            Ex.I = 0;
            Ex.Tab = new ulong [18];
            Ex.PobranieDanych(Ex.Kolumny);
            Ex.Obliczanie(Ex.Kolumny);
            Ex.Wypisanie(Ex.Kolumny, Ex.I);
        }
    }
}
