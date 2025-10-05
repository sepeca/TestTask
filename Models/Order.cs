namespace TestTask.Models;
public class Order
{
    public int OrderName { get; set; } = -1;
    public List<ProducedPieces> ProducedPieces { get; set; } = new();
}