namespace TestTask.Models;
using System.Collections.Generic;
public class Client
{
    public string ClientName { get; set; } = string.Empty;
    // Z toho důvodu, že IČ může byt např. '00001234', nebude spravně dělat to int, jinak v JSON to bude 1234, což v budoucím spravování
    //  muže byt interpretované jak '12340000' místo '00001234'
    public string ClientIC { get; set; } = string.Empty; 
    public List<Order> Orders { get; set; } = new();
}