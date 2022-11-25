using System;
using System.Collections.Generic;
using System.Text.Json;
using Newtonsoft.Json;
using System.Net;
using System.IO;
using FastExcel;
using System.Threading.Tasks;
using _LNG_Collector.Utils;

namespace _LNG_Collector
{
    internal class LNG_Situatii_Zilnice
    {
        public string settings_file = "settings.json";
        
        static void Main(string[] args)
        {
            //Pentru real-time environment
            //string data_curenta = DateTime.Now.ToString("dd-MM-yyyy");
            //string folder_curent = Directory.GetCurrentDirectory();


            //Pentru testare
            string data_curenta = "01-12-2022";
            string folder_curent = @"C:\GitHub\LNG_Collector\";

            Console.WriteLine("Initiere Generare Situatii Zilnice LNG: " + data_curenta);
            Console.WriteLine("...");
            
            //citeste fisierele de intrare
            Console.WriteLine("Verifica existenta date de intrare...");
            CitesteSiVerificaFisierele(folder_curent);

            //Creaza structura pentru data curenta
            Directory.CreateDirectory(folder_curent+@"\"+data_curenta);
            Directory.CreateDirectory(folder_curent + @"\" + data_curenta +@"\input");
            Directory.CreateDirectory(folder_curent + @"\" + data_curenta +@"\output");

            //copiaza fisierele de intrare in directorul zilnic de input
            CopiazaFisiere(folder_curent + @"\transfer_input\", folder_curent + @"\" + data_curenta + @"\input\");
            //copiaza fisierele template pentru a fi umplute cu dare in diectorul zilnic de output
            CopiazaFisiere(folder_curent + @"\templates\", folder_curent + @"\" + data_curenta + @"\output\");

            //

        }

        private static void CopiazaFisiere(string inputPath, string dailyInput)
        {
            foreach (var newPath in Directory.GetFiles(inputPath, "*.*", SearchOption.AllDirectories))
            {
                File.Copy(newPath, newPath.Replace(inputPath, dailyInput));
                Console.WriteLine(newPath+": Fisier copiat in " + dailyInput);
            }
        }

        private static void CitesteSiVerificaFisierele(string folder_curent)
        {
            //citeste din setari ce fisiere trebuie sa existe
            Setari setari = new Setari();
            try
            {
                setari = JsonConvert.DeserializeObject<Setari>(File.ReadAllText(folder_curent + @"\settings.json"));
            }catch (Exception ex)
            {
                Console.WriteLine("Fisier de setari eronat");
                return;
            }
            
            //
            foreach (string filename in setari.InputFiles)
            {
                //verify all files in the list
                bool allOK = true;
                if (File.Exists(folder_curent + @"\transfer_input\" + filename))
                {
                    Console.WriteLine(folder_curent + @"\transfer_input\" + filename + " -> exista. OK");
                }
                else { 
                    allOK = false;
                    Console.WriteLine(folder_curent + @"\transfer_input\" + filename + " -> NU EXISTA!. NOK");
                }
                if (!allOK) {
                    Console.WriteLine("Nu toate fisierele de input sunt prezente");
                    Console.WriteLine("Datele zilnice nu au fost inca incarcate. Incercati mai tarziu sau verificati folderul input din data curenta");
                    return;
                }
            }
        }

        
    }
}
