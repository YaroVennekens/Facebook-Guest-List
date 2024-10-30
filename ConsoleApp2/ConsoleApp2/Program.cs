using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

class Program
{
    static void Main()
    {
        string bestandsNaamFacebook,
               bestandsNaamExtra,
               facebookGastenlijstBestandspad,
               extraPersonenBestandspad,
               outputBestandspad;

        Console.WriteLine("Maak automatisch een gastenlijst op basis van uw geëxporteerde Facebook-gastenlijst. Deze lijst");
        Console.WriteLine("wordt alfabetisch gesorteerd en dubbele namen worden verwijderd. U kunt ook extra personen");
        Console.WriteLine("aan de gastenlijst toevoegen door een bestand met extra personen aan te maken.\n");

        bestandsNaamFacebook = ReadFileName("Geef de naam van het Facebook gastenlijst export bestand:", ".csv");
        bestandsNaamExtra = ReadFileName("Geef de naam van het extra personen bestand:", ".txt");

        facebookGastenlijstBestandspad = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"{bestandsNaamFacebook}");
        extraPersonenBestandspad = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"{bestandsNaamExtra}");
        outputBestandspad = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "gastenlijst_export.txt");

        HashSet<string> uniekeNamen = new HashSet<string>();

        try
        {
          
            if (!File.Exists(facebookGastenlijstBestandspad))
            {
                Console.WriteLine($"Het bestand '{facebookGastenlijstBestandspad}' bestaat niet.");
                return;
            }

         
            if (!File.Exists(extraPersonenBestandspad))
            {
                Console.WriteLine($"Het bestand '{extraPersonenBestandspad}' bestaat niet.");
                return; 
            }

            string[] namen = File.ReadAllLines(facebookGastenlijstBestandspad);
            foreach (string naam in namen)
            {
                string naamLowerCase = naam.ToLower();
                if (!naamLowerCase.Contains("uitgenodigd"))
                {
                    uniekeNamen.Add(naam.Trim()
                        .Replace("Aanwezig", "")
                        .Replace(",", "")
                        .Replace("\"", "")
                        .Replace("Misschien", "").Trim());
                }
            }

          
            string[] extraNamen = File.ReadAllLines(extraPersonenBestandspad);
            foreach (string naam in extraNamen)
            {
                if (!string.IsNullOrWhiteSpace(naam))
                {
                    uniekeNamen.Add(naam.Trim());
                }
            }

            List<string> gesorteerdeNamen = uniekeNamen.ToList();
            gesorteerdeNamen.Sort();

            List<string> outputNamen = new List<string>();
            char huidigeLetter = '\0';

            foreach (string orderedName in gesorteerdeNamen)
            {
                if (string.IsNullOrWhiteSpace(orderedName))
                    continue;

                char eersteLetter = char.ToUpper(orderedName[0]);

                if (eersteLetter != huidigeLetter)
                {
                    huidigeLetter = eersteLetter;
                    outputNamen.Add($"\n{eersteLetter} ----\n");
                }

                outputNamen.Add(orderedName);
            }

            
            File.WriteAllLines(outputBestandspad, outputNamen);
            foreach (string persoon in outputNamen)
            {
                Console.WriteLine(persoon);
            }

            Console.WriteLine($"Gesorteerde unieke namen zijn opgeslagen in: {outputBestandspad}. " +
                              $"Er gaan momenteel {outputNamen.Count(n => !n.Contains(" ----"))} mensen naar uw vat.");
        }
        catch (Exception e)
        {
            Console.WriteLine($"Er is een fout opgetreden: {e.Message}");
        }
    }

    static string ReadFileName(string message, string extension)
    {
        string bestandsnaam;
        while (true)
        {
            Console.Write(message);
            bestandsnaam = Console.ReadLine();

          
            if (!bestandsnaam.EndsWith(extension, StringComparison.OrdinalIgnoreCase))
            {
                bestandsnaam += extension;
            }

        
            string bestandsPad = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), bestandsnaam);
            if (File.Exists(bestandsPad))
            {
                return bestandsnaam; 
            }
            else
            {
             
                Console.WriteLine($"Het bestand '{bestandsnaam}' bestaat niet. Beschikbare bestanden op het bureaublad:");
                ShowFiles(Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            }
        }
    }

    static void ShowFiles(string directoryPath)
    {
        try
        {
            string[] files = Directory.GetFiles(directoryPath);
            foreach (string file in files)
            {
                Console.WriteLine(Path.GetFileName(file));
            }
        }
        catch (Exception e)
        {
            Console.WriteLine($"Fout bij het ophalen van bestanden: {e.Message}");
        }
    }
}
