using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

class Program
{
    static void Main()
    {
         ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Set based on your usage

    string bestandsNaamFacebook,
           bestandsNaamExtra,
           facebookGastenlijstBestandspad,
           extraPersonenBestandspad,
           outputBestandspad,
           eventNaam;

    Console.WriteLine("Maak automatisch een gastenlijst op basis van uw geëxporteerde Facebook-gastenlijst. Deze lijst");
    Console.WriteLine("wordt alfabetisch gesorteerd en dubbele namen worden verwijderd. U kunt ook extra personen");
    Console.WriteLine("aan de gastenlijst toevoegen door een bestand met extra personen aan te maken.\nZorg er voor dat uw bestanden op uw desktop staan\n");
    Console.WriteLine("Aan de gastenlijst worden alleen de personen toegevoegd die op MISSCHIEN of GAAT staan!\n");
    
    eventNaam = ReadLine("Geef de naam van uw event in: ");
    bestandsNaamFacebook = ReadFileName("Geef de naam van het Facebook gastenlijst export bestand: ", ".csv");
    bestandsNaamExtra = ReadFileName("Geef de naam van het extra personen bestand: ", ".txt");

    facebookGastenlijstBestandspad = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"{bestandsNaamFacebook}");
    extraPersonenBestandspad = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"{bestandsNaamExtra}");
    outputBestandspad = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "gastenlijst_export.txt");

    HashSet<string> uniekeNamen = new HashSet<string>();

    try
    {
        // Check if input files exist
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

        // Read names from the Facebook guest list
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

        // Read extra names
        string[] extraNamen = File.ReadAllLines(extraPersonenBestandspad);
        foreach (string naam in extraNamen)
        {
            if (!string.IsNullOrWhiteSpace(naam))
            {
                uniekeNamen.Add(naam.Trim());
            }
        }

        // Sort names and prepare for output
        List<string> gesorteerdeNamen = uniekeNamen.ToList();
        gesorteerdeNamen.Sort();

        // Create output text file
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

        Console.WriteLine($"\nGesorteerde unieke namen zijn opgeslagen in: {outputBestandspad}. " +
                          $"Er gaan momenteel {outputNamen.Count(n => !n.Contains(" ----"))} mensen naar uw evenement {eventNaam}.");

        // Create Excel file with columns for each first letter
        CreateExcelWithNameColumns(gesorteerdeNamen, eventNaam);
    }
    catch (Exception e)
    {
        Console.WriteLine($"Er is een fout opgetreden: {e.Message}");
    }
    }

    static void CreateExcelWithNameColumns(List<string> gesorteerdeNamen, string eventNaam)
    {
        using (ExcelPackage package = new ExcelPackage())
        {
            var worksheet = package.Workbook.Worksheets.Add("Gastenlijst");

            // Create a dictionary to hold names by their first letters
            Dictionary<char, List<string>> namesByFirstLetter = new Dictionary<char, List<string>>();

            foreach (var name in gesorteerdeNamen)
            {
                char firstLetter = char.ToUpper(name[0]);
                if (!namesByFirstLetter.ContainsKey(firstLetter))
                {
                    namesByFirstLetter[firstLetter] = new List<string>();
                }
                namesByFirstLetter[firstLetter].Add(name);
            }

            // Write names to columns based on their first letters
            int column = 1; // Start from the first column
            foreach (var kvp in namesByFirstLetter)
            {
                worksheet.Cells[1, column].Value = kvp.Key; // Set header as the letter
                for (int i = 0; i < kvp.Value.Count; i++)
                {
                    worksheet.Cells[i + 2, column].Value = kvp.Value[i]; // Start writing names from the second row
                }
                column++;
            }

            // Save the Excel file
            string excelFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"{eventNaam}_gastenlijst.xlsx");
            FileInfo excelFile = new FileInfo(excelFilePath);
            package.SaveAs(excelFile);

            Console.WriteLine($"Excel bestand is aangemaakt: {excelFilePath}");
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
                Console.WriteLine($"Het bestand '{bestandsnaam}' bestaat niet.\n Beschikbare bestanden op het bureaublad:");
                ShowFiles(Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            }
        }
    }

    static string ReadLine(string message)
    {
        string naam;
        do
        {
            Console.Write(message);
            naam = Console.ReadLine();
        } while (string.IsNullOrWhiteSpace(naam));

        return naam;
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
