namespace TestTask.Models;
public class Order
{
    public int OrderName { get; set; }
    public List<ProducedPieces> ProducedPieces { get; set; } = new();
}