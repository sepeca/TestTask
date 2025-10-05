using Newtonsoft.Json;
//Vyhledavani xlsx file
var projectDir = Directory.GetCurrentDirectory();

var excelFile = Directory.GetFiles(projectDir, "*.xlsx", SearchOption.TopDirectoryOnly).FirstOrDefault();
if (excelFile == null)
{
    Console.WriteLine("No .xlsx file was found");
    return;
}
// Vytahovani klientu z tabulky
var clients = ExcelReader.Read(excelFile);
// Data do JSONu
var json = JsonConvert.SerializeObject(clients, Formatting.Indented);
// Export do file
File.WriteAllText("output.json", json);

Console.WriteLine("Data is successfully stored in output.json");