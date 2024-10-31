using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

class Program
{
    static void Main()
    {
         ExcelPackage.LicenseContext = LicenseContext.NonCommercial; 

    string fileFacebook,
           fileExtraPersons,
           facebookGuestlistFilePath,
           extraPersonFilePath,
           outputFilePath,
           eventName;
    
    Console.WriteLine("Maak automatisch een gastenlijst op basis van uw geëxporteerde Facebook-gastenlijst. Deze lijst");
    Console.WriteLine("wordt alfabetisch gesorteerd en dubbele names worden verwijderd. U kunt ook extra personen");
    Console.WriteLine("aan de gastenlijst toevoegen door een bestand met extra personen aan te maken.\nZorg er voor dat uw bestanden op uw desktop staan\n");
    Console.WriteLine("Aan de gastenlijst worden alleen de personen toegevoegd die op MISSCHIEN of GAAT staan!\n");
    
    eventName = ReadLine("Geef de name van uw event in: ");
    fileFacebook = ReadFileName("Geef de name van het Facebook gastenlijst export bestand (.csv): ", ".csv");
    fileExtraPersons = ReadFileName("Geef de name van het extra personen bestand (.txt): ", ".txt");
    
    facebookGuestlistFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"{fileFacebook}");
    extraPersonFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"{fileExtraPersons}");
    outputFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "gastenlijst_export.txt");

    HashSet<string> uniqueNames = new HashSet<string>();

    try
    {
      
        if (!File.Exists(facebookGuestlistFilePath))
        {
            Console.WriteLine($"Het bestand '{facebookGuestlistFilePath}' bestaat niet.");
            return;
        }

        if (!File.Exists(extraPersonFilePath))
        {
            Console.WriteLine($"Het bestand '{extraPersonFilePath}' bestaat niet.");
            return; 
        }

       
        string[] names = File.ReadAllLines(facebookGuestlistFilePath);
        foreach (string name in names)
        {
            string nameLowerCase = name.ToLower();
            if (!nameLowerCase.Contains("uitgenodigd"))
            {
                uniqueNames.Add(name.Trim()
                    .Replace("Aanwezig", "")
                    .Replace(",", "")
                    .Replace("\"", "")
                    .Replace("Misschien", "").Trim());
            }
        }
        
        string[] extranames = File.ReadAllLines(extraPersonFilePath);
        foreach (string name in extranames)
        {
            if (!string.IsNullOrWhiteSpace(name))
            {
                uniqueNames.Add(name.Trim());
            }
        }
       
        List<string> sortedNames = uniqueNames.ToList();
        sortedNames.Sort();
       
        List<string> outputnames = new List<string>();
        char currentLetter = '\0';

        foreach (string sortedName in sortedNames)
        {
            if (string.IsNullOrWhiteSpace(sortedName))
                continue;

            char firstLetter = char.ToUpper(sortedName[0]);

            if (firstLetter != currentLetter)
            {
                currentLetter = firstLetter;
                outputnames.Add($"\n{firstLetter} ----\n");
            }

            outputnames.Add(sortedName);
        }

        File.WriteAllLines(outputFilePath, outputnames);
        int amount = 1;
        
        foreach (string persoon in outputnames)
        {
            amount++;
            Console.WriteLine(persoon);
        }

        Console.WriteLine($"\nGesorteerde unieke names zijn opgeslagen in: {outputFilePath}. " +
                          $"Er gaan momenteel {amount} mensen naar uw evenement {eventName}.");

        CreateExcelWithNameColumns(sortedNames, eventName);
    }
    catch (Exception e)
    {
        Console.WriteLine($"Er is een fout opgetreden: {e.Message}");
    }
    }
    static void CreateExcelWithNameColumns(List<string> sortedNames, string eventName)
    {
        using (ExcelPackage package = new ExcelPackage())
        {
            var worksheet = package.Workbook.Worksheets.Add("Gastenlijst");

            Dictionary<char, List<string>> namesByFirstLetter = new Dictionary<char, List<string>>();

            foreach (var name in sortedNames)
            {
                char firstLetter = char.ToUpper(name[0]);
                if (!namesByFirstLetter.ContainsKey(firstLetter))
                {
                    namesByFirstLetter[firstLetter] = new List<string>();
                }
                namesByFirstLetter[firstLetter].Add(name);
            }

            int column = 1; 
            foreach (var kvp in namesByFirstLetter)
            {
                worksheet.Cells[1, column].Value = kvp.Key; 
                for (int i = 0; i < kvp.Value.Count; i++)
                {
                    worksheet.Cells[i + 2, column].Value = kvp.Value[i]; 
                }

             
                worksheet.Column(column).Width = 40;

            
                for (int i = 1; i <= kvp.Value.Count + 1; i++) 
                {
                    worksheet.Cells[i, column].Style.Font.Size = 18; 
                    worksheet.Row(i).Height = 20;
                }

                column++;
            }

            string excelFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"{eventName}_gastenlijst.xlsx");
            FileInfo excelFile = new FileInfo(excelFilePath);
            package.SaveAs(excelFile);

            Console.WriteLine($"Excel bestand is aangemaakt: {excelFilePath}");
        }
    }

    static string ReadFileName(string message, string extension)
    {
        string bestandsname;
        while (true)
        {
            Console.Write(message);
            bestandsname = Console.ReadLine();

            if (!bestandsname.EndsWith(extension, StringComparison.OrdinalIgnoreCase))
            {
                bestandsname += extension;
            }

            string bestandsPad = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), bestandsname);
            if (File.Exists(bestandsPad))
            {
                return bestandsname; 
            }
            else
            {
                Console.WriteLine($"Het bestand '{bestandsname}' bestaat niet.\n Beschikbare bestanden op het bureaublad:");
                ShowFiles(Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            }
        }
    }

    static string ReadLine(string message)
    {
        string name;
        do
        {
            Console.Write(message);
            name = Console.ReadLine();
        } while (string.IsNullOrWhiteSpace(name));

        return name;
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
