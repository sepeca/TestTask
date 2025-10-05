using ClosedXML.Excel;
using TestTask.Models;

public static class ExcelReader
{
    public static List<Client> Read(string filePath)
    {
        // Vytvoreni dictionary pro efektivnejsi vyhledavani existujicich klientu
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
        foreach (var row in sheet.RowsUsed().Skip(1))
        {
            // Parse cislo poprve a overeni jestli klinta s cislem uz mame
            var clientIC = row.Cell(2).GetValue<string>();
            
            if (!clientsDict.TryGetValue(clientIC, out var client))
            {
                // Pokud takoveho klienta nemame, tak uz dava smysl odebirat jmeno
                var clientName = row.Cell(1).GetValue<string>();
                client = new Client { ClientName = clientName, ClientIC = clientIC };
                clientsDict[clientIC] = client;
            }
            // Vytvoreni zakazky s podrobnostmi
            var orderName = row.Cell(3).GetValue<int>();
            var order = new Order { OrderName = orderName };

            for (int i = 0; i < month.Length; i++)
            {
                var cell = row.Cell(4 + i);
                var numPieces = cell.GetValue<int>();
                order.ProducedPieces.Add(new ProducedPieces { Period = month[i], NumPieces = numPieces });
            }
            // Prirazeni zakazky
            client.Orders.Add(order);
        }

    return clientsDict.Values.ToList();
    }
}
