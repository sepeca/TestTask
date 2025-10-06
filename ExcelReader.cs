using ClosedXML.Excel;
using TestTask.Models;

public static class ExcelReader
{
    private static void ErrorWriter(int rowNumber, int columnNumber, string value, string place)
    { 
        Console.WriteLine($"Error in raw {rowNumber}, column {columnNumber}: value {value} is not a number.");
        Console.WriteLine($"Problem place is {place}. Used placeholder -1");
    }
    public static List<Client> Read(string filePath)
    {
        // Vytvoreni dictionary pro efektivnejsi vyhledavani existujicich
        var clientsDict = new Dictionary<string, Client>();

        // Prevzeti xlsx
        using var workbook = new XLWorkbook(filePath);
        var sheet = workbook.Worksheets.First();

        //Hledani zacatecneho roku v tabulce
        var header = sheet.Cell(1, 4).GetValue<string>();
        int startYear;

        // Snaha parsingu pokud format je rrrr-mm
        if (!int.TryParse(header[..4], out startYear))
        {
            // Snaha parsingu pokud format je mm-rrrr
            if (!int.TryParse(header[5..9], out startYear))
            {
                throw new Exception($"Unable to recognize year from 4th header: {header}");
            }
        }
        // Pole mesice (D-AA)
        var month = new string[24];
        for (int y = 0; y < 2; y++)
            for (int m = 0; m < 12; m++)
                month[y * 12 + m] = $"{startYear + y}-{(m + 1):D2}";

        // Prochazime radky, zaciname s 2.
        foreach (IXLRow row in sheet.RowsUsed().Skip(1))
        {
            // Parse IC
            string clientIC = row.Cell(2).GetValue<string>();

            // Overeni jestli cislo je spravne
            int.TryParse(clientIC, out var num);
            if (num < 1 || num > 99999999) { throw new Exception($"IC is corrupted on {row.RowNumber()}th row"); }

            //Overeni jestli klient uz nebyl vypsan
            if (!clientsDict.TryGetValue(clientIC, out var client))
            {
                // Pokud takoveho klienta nemame, tak uz dava smysl odebirat jmeno
                var clientName = row.Cell(1).GetValue<string>();
                client = new Client { ClientName = clientName, ClientIC = clientIC };
                clientsDict[clientIC] = client;
            }
            // Vytvoreni zakazky s podrobnostmi
            int orderName;
            
            //Zpracovani chyby ve pripade, ze je prazdna/spatna hodnota v bunce 
            try { orderName = row.Cell(3).GetValue<int>(); }
            catch
            {
                orderName = -1;
                ErrorWriter(row.RowNumber(), 3, row.Cell(3).GetValue<string>(), "order name");

            }
            var order = new Order { OrderName = orderName };

            for (int i = 0; i < month.Length; i++)
            {
                var cell = row.Cell(4 + i);
                int numPieces;

                //Zpracovani chyby ve pripade, ze je prazdna/spatna hodnota v bunce 
                try { numPieces = cell.GetValue<int>(); }
                catch
                {
                    numPieces = -1;
                    ErrorWriter(row.RowNumber(), 4 + i, cell.GetValue<string>(), "quantity of pieces");
                }
                order.ProducedPieces.Add(new ProducedPieces { Period = month[i], NumPieces = numPieces });
            }
            // Prirazeni zakazky
            client.Orders.Add(order);
        }

        return clientsDict.Values.ToList();
    }
}
