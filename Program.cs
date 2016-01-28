using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelObj
{
    class Excel
    {
         ulong Kolumny, I;
         ulong[] Tab;


         void Zmiana_Danych(ulong kolumny, ulong i)
        {
            Kolumny = kolumny;
            I = i;
        }

         void Zmiana_Tablicy(ulong a)
        {
            Tab[I] = a;
            I++;
        }
         ulong Czytanie_Tablicy(int i)
        {
            return Tab[i];

        }
               
         void PobranieDanych()
        {
            ulong kolumny= 0;

            Console.WriteLine("\nProsze podac o ktora kolumne chodzi: \n");
                 try
                    {
                        kolumny = Convert.ToUInt64(Console.ReadLine());
                    }
                 catch (Exception ex) 
                    { 
                        Console.WriteLine
                         ("Gratulacje, zhakowales system, \nwpisales litery, za duzo cyfr, \nczy nic?\n");
                    }
            if (kolumny > 18446744073709551614)
                {
                    Console.WriteLine("Przepraszam najwieksza liczba jest 18 446 744 073 709 551 615");
                    Console.ReadLine();
                    Zmiana_Danych(18446744073709551614, 0);
                }
                else
                   Zmiana_Danych(kolumny, 0);
        }
        
         void Obliczanie(ulong kolumny)
        {
            while (kolumny > 0)
            {
                Zmiana_Tablicy(kolumny % 26);
                kolumny = kolumny / 26;
            }
        }
        
         void Wypisanie(ulong kolumny, ulong i)
        {
             Console.Write("Szukasz kolumny: ");
             int a = Convert.ToInt32(i);
            for (int j = (a - 1); j >= 0; j--)
            {
                if (Czytanie_Tablicy(j) > 0)
                    Console.Write((char)(Czytanie_Tablicy(j) + 64));
                else
                    Console.Write(Czytanie_Tablicy(j));
            }
            
            Console.WriteLine("\n");
           
            //JeszczeRaz();
         }

         void JeszczeRaz()
        {
            Console.WriteLine("Jeszcze raz? Spacja - Tak \n");

            ConsoleKeyInfo info = Console.ReadKey();
            if (info.Key == ConsoleKey.Spacebar)
            {
                PobranieDanych();
                Obliczanie(Kolumny);
                Wypisanie(Kolumny, I);
                JeszczeRaz();
            }
            else
                Console.WriteLine("\nDziekuje");
            Console.ReadLine();
        }
        
      // Tutaj zaczyna sie program 
 
    static void Main(string[] args)
        {

            Excel Ex = new Excel();
            Ex.Kolumny = 0;
            Ex.I = 0;
            Ex.Tab = new ulong [18];

            Ex.PobranieDanych();
            Ex.Obliczanie(Ex.Kolumny);
            Ex.Wypisanie(Ex.Kolumny, Ex.I);
            Ex.JeszczeRaz();
        }
    }
}
