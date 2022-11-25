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
        string settings_file = "settings.json";
        
        static void Main(string[] args)
        {
            string data_curenta = DateTime.Now.ToString("dd/MM/yyyy");

            Console.WriteLine("Initiere Generare Situatii Zilnice LNG: " + data_curenta);
            Console.WriteLine("...");
            
            //citeste fisierele de intrare
            Console.WriteLine("Verifica existenta date de intrare...");
            CitesteSiVerificaFisierele();

            //Creaza structura pentru data curenta
            Directory.CreateDirectory(data_curenta);
            Directory.CreateDirectory(data_curenta+@"\input");
            Directory.CreateDirectory(data_curenta+@"\output");

            //copiaza fisierele de intrare in directorul zilnic de input
            CopiazaFisiere(@"\transfer_input\", data_curenta + @"\input\");
            //copiaza fisierele template pentru a fi umplute cu dare in diectorul zilnic de output
            CopiazaFisiere(@"\templates\", data_curenta + @"\output\");
        }

        private static void CopiazaFisiere(string inputPath, string dailyInput)
        {
            foreach (var newPath in Directory.GetFiles(inputPath, "*.*", SearchOption.AllDirectories))
            {
                File.Copy(newPath, newPath.Replace(inputPath, dailyInput));
            }
        }

        private static void CitesteSiVerificaFisierele()
        {
            //citeste din setari ce fisiere trebuie sa existe
            Setari setari = JsonConvert.DeserializeObject<Setari>(File.ReadAllText("settings.json"));
            if (setari == null) {
                Console.WriteLine("Fisier de setari eronat");
                return;
            }

            //
            foreach (string filename in setari.input_files)
            {
                //verify all files in the list
                bool allOK = true;
                if (File.Exists(@"\transfer_input\"+filename))
                {
                    Console.WriteLine(@"\transfer_input\"+ filename + "-> exista. OK");
                }
                else { 
                    allOK = false;
                    Console.WriteLine(@"\transfer_input\" + filename + " -> NU EXISTA!. NOK");
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
