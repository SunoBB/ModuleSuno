
using System.ComponentModel;
using System.Diagnostics;
using OfficeOpenXml;

// EPPLUS 
public class Item
{
    public int STT { get; set; }
    public string ItemID { get; set; } // Mã định danh duy nhất cho mỗi mặt hàng (kiểu string).
    public string ItemName { get; set; } // Tên của mặt hàng (kiểu chuỗi).
    public string Description { get; set; } // Mô tả chi tiết về mặt hàng (kiểu chuỗi hoặc văn bản).
    public decimal Price { get; set; } // Giá của mặt hàng (kiểu số thập phân).
    public int Quantity { get; set; } // Số lượng hiện có trong kho (kiểu số nguyên).
    public DateTime DateAdded { get; set; } // Ngày mặt hàng được thêm vào kho (kiểu ngày tháng).
    public string Supplier { get; set; } // Thông tin về nhà cung cấp của mặt hàng (kiểu chuỗi).


    public static void AddItem(List<Item> itemList, Item newItem)
    {
        newItem.STT = itemList.Count + 1;
        itemList.Add(newItem);
        Console.WriteLine($"Item '{newItem.ItemName}' added!");
    }


    public static void EditItem(List<Item> itemList, string itemId, Item updatedItem)
    {
        // ItemID, ItemName, Description, Price, Quantity, DateAdded, Supplier

        // Console.WriteLine("Updating");

        Item findItemEdit = itemList.FirstOrDefault(item => itemId == item.ItemID);

        if (findItemEdit != null)
        {
            findItemEdit.ItemName = updatedItem.ItemName;
            findItemEdit.Description = updatedItem.Description;
            findItemEdit.Price = updatedItem.Price;
            findItemEdit.Quantity = updatedItem.Quantity;
            findItemEdit.DateAdded = updatedItem.DateAdded;
            findItemEdit.Supplier = updatedItem.Supplier;

            Console.WriteLine($"ID Item: {itemId} was edited!");
        }
        else
        {
            Console.WriteLine($"ItemID: {itemId} not found!");
        }
    }

    public static void RemoveItem(List<Item> itemList, string itemId)
    {
        Item intemToRemove = itemList.FirstOrDefault(item => item.ItemID == itemId);

        if (intemToRemove != null)
        {
            itemList.Remove(intemToRemove);
            Console.WriteLine($"Item with ID {itemId} was del");

        }
        else
        {
            Console.WriteLine($"Item ID not found");
        }
    }

    public static void SearchItem(List<Item> itemList, string itemName)
    {
        Item foundItem = itemList.FirstOrDefault(item => item.ItemName == itemName);

        if (foundItem != null)
        {
            Console.WriteLine(foundItem.ItemID);
            Console.WriteLine("more..."); //More ...

        } 
        else
        {
            Console.WriteLine("Item not found!");
        }
        
    }

    public static void ShowItem(List<Item> itemList)
    {
        foreach (Item item in itemList)
        {
            // ItemID, ItemName, Description, Price, Quantity, DateAdded, Supplier
            Console.WriteLine($" {item.ItemID} {item.ItemName}");
        }
    }
}

public class Program
{
    static void Main()
    {
        // Add, Search, Del, Edit, Exit
        List<Item> itemList = new List<Item>();
        int choice;

        do
        {
            Console.WriteLine("0. Exit \n1. ADD Item \n2. Search Item \n3. RemoveItem \n4. Edit Item \n5.Show ");

            do
            {
                if (int.TryParse(Console.ReadLine(), out choice) && choice >= 0 && choice <= 5)
                {
                    break;
                } 
                else
                {
                    Console.WriteLine("ReChoice!");
                }
            } while (true);

            if (choice == 0)
            {
                break;
            } 
            else if (choice == 1) // ADD
            {
                Item.AddItem(itemList, new Item
                {
                    // STT = 1,
                    ItemID = "01005",
                    ItemName = "Laptop",
                    Description = "Powerful laptop",
                    Price = 1200.50m,
                    Quantity = 10,
                    DateAdded = DateTime.Now,
                    Supplier = "ABC Electronics"
                });
            }
            else if (choice == 2) // Search
            {
                Item.SearchItem(itemList, "itemName");
            }
            else if (choice == 3) // Remove
            {
                Item.RemoveItem(itemList, "ItemID");
            }
            else if (choice == 4) // Edit
            {
                Item updatedItem = new Item
                {
                    ItemName = "",
                    Description = "",
                    Price = 0.1m,
                    Quantity = 10,
                    DateAdded = DateTime.Now,
                    Supplier = ""

                };
                Item.EditItem(itemList, "iemID", updatedItem);
            }
            else if (choice == 5)
            {
                Item.ShowItem(itemList);
            }
            {

            }
        } while (true);
    }
}


/*class Program
{
    static void Main()
    {
        List<Item> itemList = new List<Item>();

        Item.AddItem(itemList, new Item
        {
            // STT = 1,
            ItemID = "01005",
            ItemName = "Laptop",
            Description = "Powerful laptop",
            Price = 1200.50m,
            Quantity = 10,
            DateAdded = DateTime.Now,
            Supplier = "ABC Electronics"
        });


        // Retrun Result Search Item
        Item.SearchItem(itemList, "iTemName");

        Item.RemoveItem(itemList, "itemID");


        // ItemName, Description, Price, Quantity, DateAdded, Supplier
        Item updatedItem = new Item
        {
            ItemName = "",
            Description = "",
            Price = 0.1m,
            Quantity = 10,
            DateAdded = DateTime.Now,
            Supplier = ""

        };
        Item.EditItem(itemList, "itemID", updatedItem);
    }
}*/
