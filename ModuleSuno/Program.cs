
using OfficeOpenXml; // EPPLUS 



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
}

public class _Operator
{
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
            Console.WriteLine($"> {item.STT} {item.ItemID} {item.ItemName} {item.Description} {item.Price} {item.Quantity} {item.DateAdded} {item.Supplier}");
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

public class _Package
{

    public static void Noti()
    {
        Console.WriteLine("\t\t\t----------Menu---------- ");
        Console.WriteLine("0. Exit \n1. ADD Item \n2. Search Item \n3. RemoveItem \n4. Edit Item \n5. Show Item\n6. Export\n");

    }

    public static void ImportDt(List<Item> itemList, string filePath)
    {
        try
        {
            if (File.Exists(filePath)) // Check file exist
            {
                using (var pkg = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet ws = pkg.Workbook.Worksheets["Persons"];

                    int rowCount = ws.Dimension.Rows;

                    //Start with row 2 to skip the header
                    for (int row = 2; row <= rowCount; row++)
                    {
                        double _DateAdded = double.Parse(ws.Cells[row, 6].Text);
                        DateTime DateAdd = DateTime.FromOADate(_DateAdded);

                        // Console.WriteLine($"{ws.Cells[row, col].Text}\t");
                        // ADD Item into List
                        Item ToImportItem = new Item
                        {

                            ItemID = ws.Cells[row, 1].Text,
                            ItemName = ws.Cells[row, 2].Text,
                            Description = ws.Cells[row, 3].Text,
                            Price = decimal.Parse(ws.Cells[row, 4].Text),
                            Quantity = int.Parse(ws.Cells[row, 5].Text),
                            DateAdded = DateAdd,
                            Supplier = ws.Cells[row, 7].Text

                        };

                        // System.Console.WriteLine($"{ToImportItem.ItemID} {ToImportItem.ItemName} {ToImportItem.Description} {ToImportItem.Price} {ToImportItem.Quantity} {ToImportItem.DateAdded} {ToImportItem.Supplier}");
                        _Operator.AddItem(itemList, ToImportItem);
                        System.Console.WriteLine("Importing...");


                    }
                    Console.WriteLine("Success Import Excel File!");
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



public class Program
{
    static void Main()
    {
        // Add, Search, Del, Edit, Exit
        List<Item> itemList = new List<Item>();
        string filePath = @"C:\Users\suno\OneDrive - UET\Proj\DB\ManagerDB.xlsx";
        int choice;

        // Auto Import Data from Excel

        _Package.ImportDt(itemList, filePath);
        // show STT



        do
        {
            _Package.Noti();
            Console.Write(" >  ");
            do
            {
                if (int.TryParse(Console.ReadLine(), out choice) && choice >= 0 && choice <= 6) // Check input (Choice) in range 0, 5
                {
                    break;
                }
                else
                {
                    _Package.Noti();
                    Console.Write("Please enter a number between 0 and 6: ");
                }
            } while (true);

            if (choice == 0)
            {
                break;
            }
            else if (choice == 1) // ADD
            {
                _Operator.AddItem(itemList, _Operator.InputItem());
            }
            else if (choice == 2) // Search
            {
                Console.Write("ItemName Search: ");
                string? SearchitemName = Console.ReadLine();
                _Operator.SearchItem(itemList, SearchitemName);
            }
            else if (choice == 3) // Remove
            {
                Console.Write("ItemID Remove: ");
                string? RemoveItemID = Console.ReadLine();
                _Operator.RemoveItem(itemList, RemoveItemID);
            }
            else if (choice == 4) // Edit
            {
                Item updatedItem = _Operator.InputItem();
                _Operator.EditItem(itemList, updatedItem.ItemID, updatedItem);
            }
            else if (choice == 5)
            {
                _Operator.ShowItem(itemList);
            }

            else if (choice == 6) //Export
            {
                Console.WriteLine("Export data");
                _Package.ExportDt(filePath, itemList);
                break;
            }
            // else if (choice == 7) //Import
            // {
            //     _Package.ImportDt(filePath);
            // }
        } while (true);
    }
}
