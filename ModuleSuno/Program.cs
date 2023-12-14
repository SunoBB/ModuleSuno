
using System.ComponentModel;
using System.Diagnostics;
using OfficeOpenXml;

// EPPLUS 

public class _Package
{

    public static void Noti()
    {
        Console.WriteLine("\t\t\t----------Menu---------- ");
        Console.WriteLine("0. Exit \n1. ADD Item \n2. Search Item \n3. RemoveItem \n4. Edit Item \n5. Show Item\n6. Export\n7. Import");

    }

    public static void ImportDt(string filePath)
    {
        try
        {
            if (File.Exists(filePath))
            {
                using (var pkg = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet ws = pkg.Workbook.Worksheets["Persons"];

                    int rowCount = ws.Dimension.Rows;
                    int columnCount = ws.Dimension.Columns;

                    for (int row = 1;  row <= rowCount; row++) 
                    {
                        for (int col  = 1; col <= columnCount; col++)
                        {
                            Console.WriteLine($"{ws.Cells[row, col].Text}\t");
                        }
                        Console.WriteLine();
                    }
                }

            }
            else
            {
                Console.WriteLine("Error Code/File Not Found!");
            }
        }

        catch (Exception ex) 
        {
            Console.WriteLine(ex.Message);
        }
    }

    public static void ExportDt(string filePath, List<Item> LstItem)
    {
        try
        {
            using (var pkg = new ExcelPackage(new FileInfo(filePath)))
            {
                var ws = pkg.Workbook.Worksheets["Persons"];

                int indexRow = ws.Dimension.Rows; // Checking line in Excel

                indexRow += 1; // Tránh bị đè DL trong Excel 

                foreach (Item item in LstItem)
                {
                    Console.WriteLine("Adding....");
                    ws.Cells[indexRow, 1].Value = item.ItemID;
                    ws.Cells[indexRow, 2].Value = item.ItemName;
                    ws.Cells[indexRow, 3].Value = item.Description;
                    ws.Cells[indexRow, 4].Value = item.Price;
                    ws.Cells[indexRow, 5].Value = item.Quantity;
                    ws.Cells[indexRow, 6].Value = item.DateAdded;
                    ws.Cells[indexRow, 7].Value = item.Supplier;

                    indexRow++;
                    pkg.Save();
                }
            }

        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
    }
}

public class Item
{
    public int STT { get; set; }
    public string? ItemID { get; set; } // Mã định danh duy nhất cho mỗi mặt hàng (kiểu string).
    public string? ItemName { get; set; } // Tên của mặt hàng (kiểu chuỗi).
    public string? Description { get; set; } // Mô tả chi tiết về mặt hàng (kiểu chuỗi hoặc văn bản).
    public decimal? Price { get; set; } // Giá của mặt hàng (kiểu số thập phân).
    public int Quantity { get; set; } // Số lượng hiện có trong kho (kiểu số nguyên).
    public DateTime DateAdded { get; set; } // Ngày mặt hàng được thêm vào kho (kiểu ngày tháng).
    public string? Supplier { get; set; } // Thông tin về nhà cung cấp của mặt hàng (kiểu chuỗi).


    public static void AddItem(List<Item> itemList, Item newItem)
    {
        newItem.STT = itemList.Count + 1;
        itemList.Add(newItem);
        Console.WriteLine($"Item '{newItem.ItemName}' added!");
    }

    public static void EditItem(List<Item> itemList, string itemId, Item updatedItem)
    {
        // ItemID, ItemName, Description, Price, Quantity, DateAdded, Supplier

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
            Console.WriteLine("Item found! \t");
            Console.WriteLine(foundItem.ItemID);
            Console.WriteLine(foundItem.ItemName);
            Console.WriteLine(foundItem.Description);
            Console.WriteLine(foundItem.Price);
            Console.WriteLine(foundItem.Quantity);
            Console.WriteLine(foundItem.DateAdded);
            Console.WriteLine(foundItem.Supplier);
            Console.Write("\t");
        }
        else
        {
            Console.WriteLine("Item not found!");
        }

    }

    public static void ShowItem(List<Item> itemList) // Upate to Export file excel
    {
        foreach (Item item in itemList)
        {
            // ItemID, ItemName, Description, Price, Quantity, DateAdded, Supplier
            Console.WriteLine($"> {item.ItemID} {item.ItemName} {item.Description} {item.Price} {item.Quantity} {item.DateAdded} {item.Supplier}");
        }
    }

    // 
    public static Item InputItem()
    {

        Console.Write("ItemID: ");
        string? _itemID = Console.ReadLine();
        Console.Write("ItemName: ");
        string? _itemName = Console.ReadLine();
        Console.Write("Description: ");
        string? _description = Console.ReadLine();
        Console.Write("Price: ");
        decimal? _price = decimal.Parse(Console.ReadLine());
        Console.Write("Quantity: ");
        int _quantity = int.Parse(Console.ReadLine());
        DateTime _DateAdded = DateTime.Now;
        Console.Write("Supplier: ");
        string? _supplier = Console.ReadLine();

        return new Item
        {
            ItemID = _itemID,
            ItemName = _itemName,
            Description = _description,
            Price = _price,
            Quantity = _quantity,
            DateAdded = _DateAdded,
            Supplier = _supplier
        };
    }
}


public class Program
{
    static void Main()
    {
        // Add, Search, Del, Edit, Exit
        List<Item> itemList = new List<Item>();
        string filePath = @"C:\Users\suno\OneDrive - UET\Proj\DB\ManagerDB.xlsx";
        int choice;

        do
        {
            _Package.Noti();
            Console.Write(" >  ");
            do
            {
                if (int.TryParse(Console.ReadLine(), out choice) && choice >= 0 && choice <= 7) // Check input (Choice) in range 0, 5
                {
                    break;
                }
                else
                {
                    _Package.Noti();
                    Console.Write("Please enter a number between 0 and 5: ");
                }
            } while (true);

            if (choice == 0)
            {
                break;
            }
            else if (choice == 1) // ADD
            {
                Item.AddItem(itemList, Item.InputItem());
            }
            else if (choice == 2) // Search
            {
                Console.Write("ItemName Search: ");
                string? SearchitemName = Console.ReadLine();
                Item.SearchItem(itemList, SearchitemName);
            }
            else if (choice == 3) // Remove
            {
                Console.Write("ItemID Remove: ");
                string? RemoveItemID = Console.ReadLine();
                Item.RemoveItem(itemList, RemoveItemID);
            }
            else if (choice == 4) // Edit
            {
                Item updatedItem = Item.InputItem();
                Item.EditItem(itemList, updatedItem.ItemID, updatedItem);
            }
            else if (choice == 5)
            {
                Item.ShowItem(itemList);
            }

            else if (choice == 6) //Export
            {
                Console.WriteLine("Export data");
                _Package.ExportDt(filePath, itemList);
                break;
            }
            else if (choice == 7) //Import
            {
                _Package.ImportDt(filePath);
                break;
            }
        } while (true);
    }
}
