using System;
using System.CodeDom;
//using Microsoft.Office.Interop.Excel;


// Main attributes class
public class Account
{
    public string AccountNumber { get; }
    public string AccountHolder { get; set; }
    public decimal Balance { get; protected set; }

    public Account(string accountNumber, string accountHolder, decimal initialBalance)
    {
        AccountNumber = accountNumber;
        AccountHolder = accountHolder;
        Balance = initialBalance;
    }

    public virtual void Deposit(decimal amount)
    {
        if (amount > 0)
        {
            Balance += amount;
            Console.WriteLine($"\n{amount:C} deposited into account {AccountNumber}. New balance: {Balance:C}");
        }
        else
        {
            Console.WriteLine("Invalid deposit amount.");
        }
    }

    public virtual void Withdraw(decimal amount)
    {
        if (amount > 0 && Balance >= amount)
        {
            Balance -= amount;
            Console.WriteLine($"\n{amount:C} withdrawn from account {AccountNumber}. New balance: {Balance:C}");
        }
        else
        {
            Console.WriteLine("Insufficient funds or invalid withdrawal amount.");
        }
    }

    public virtual void DisplayAccountInfo()
    {
        Console.WriteLine($"Account Number: {AccountNumber}");
        Console.WriteLine($"Account Holder: {AccountHolder}");
        Console.WriteLine($"Balance: {Balance:C}");
    }
}

// Conta poupança
public class SavingsAccount : Account
{
    private decimal _interestRate;

    public SavingsAccount(string accountNumber, string accountHolder, decimal initialBalance, decimal interestRate)
        : base(accountNumber, accountHolder, initialBalance)
    {
        _interestRate = interestRate;
    }

    public void AddInterest()
    {
        decimal interestAmount = Balance * _interestRate;
        Balance += interestAmount;
        Console.WriteLine($"Interest of {interestAmount:C} added. New balance: {Balance:C}");
    }
}

// Conta Corrente
public class CheckingAccount : Account
{
    private decimal _overdraftLimit;
    // accountNumber = ID | accountHolder = Account Owner | overdraftLimit = SOmething like a special bill
    public CheckingAccount(string accountNumber, string accountHolder, decimal initialBalance, decimal overdraftLimit)
        : base(accountNumber, accountHolder, initialBalance)
    {
        _overdraftLimit = overdraftLimit;
    }

    public override void Withdraw(decimal amount)
    {
        if (amount > 0 && (Balance + _overdraftLimit) >= amount)
        {
            Balance -= amount;
            Console.WriteLine($"{amount:C} withdrawn from account {AccountNumber}. New balance: {Balance:C}");
        }
        else
        {
            Console.WriteLine("Insufficient funds or invalid withdrawal amount.");
        }
    }
}

// Add data into a Excel File
//class HandleData
//{
//    public static void readExcel()
//    {

//    }

//    public static void writeExcel()
//    {
//        string filePath = "C:\\Code\\CSharp\\data";
//        Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
//        Workbook wb;
//        Worksheet ws;

//        wb =  excel.Workbooks.Open(filePath);
//        ws = wb.Worksheets[1];

//        // AccountID | AccountOwner | Balance | OverDraftLimit
//        Range cellRange = ws.Range["A1:D1"];

//        wb.SaveAs(filePath);
//        wb.Close();
//    }
//}


class Program
{
    static void Main(string[] args)
    {
        // Creating a Checking Account
        Console.WriteLine("Input the ID");
        string accountkey = Console.ReadLine();

        Console.WriteLine("Input the name of the Account Owner: ");
        string accountOwner = Console.ReadLine();

        Console.WriteLine("Input the Initial Balance: ");
        decimal initialBalance = decimal.Parse(Console.ReadLine());

        Console.WriteLine("Input the Over Draft Limit: ");
        decimal overDraftLimit = decimal.Parse(Console.ReadLine());

        CheckingAccount checkingAccount = new CheckingAccount(accountkey, accountOwner, initialBalance, overDraftLimit);

        Console.WriteLine($"You've crated a Account with ID: {checkingAccount.AccountNumber}\nThe Accunt Owner is: {checkingAccount.AccountHolder}");

        // Creating a key to turn off program when the user quit.
        bool _switch = true;
        int option;

        while (_switch == true)
        {

            //Console.WriteLine("\nSelect a Option: (1 - 5) \n");
            //Console.WriteLine("1. Withdraw\n2. Deposit\n3. View Account Infos\n4. Add Interest\n5. Exit Program\n");
            //int option = int.Parse(Console.ReadLine());

            //while (option < 1 && option > 5)
            //{
            //    Console.WriteLine("\nInvalid option. Please, select another one: \n");
            //    option = int.Parse(Console.ReadLine());
            //}

            do
            {
                Console.WriteLine("\nSelect a Option: (1 - 5) \n");
                Console.WriteLine("1. Withdraw\n2. Deposit\n3. View Account Infos\n4. Add Interest\n5. Exit Program\n");
                option = int.Parse(Console.ReadLine());
            } while (option < 1 && option > 5);

            switch (option)
            {
                case 1:  // Withdraw
                    Console.WriteLine("\nHow much money do you want to Withdraw? \n");
                    decimal withdrawAmount = decimal.Parse(Console.ReadLine());

                    checkingAccount.Withdraw(withdrawAmount);
                    Console.WriteLine("\nFinish Operation (Press Enter)\n---------------------------------------");
                    Console.ReadKey();
                    break;
                case 2:  // Deposit
                    Console.WriteLine("\nHow much money do you want to Deposit? \n");
                    decimal depositAmount = decimal.Parse(Console.ReadLine());

                    checkingAccount.Deposit(depositAmount);
                    Console.WriteLine("\nFinish Operation (Press Enter)\n---------------------------------------");
                    Console.ReadKey();
                    break;
                case 3:  // DisplayAccountInfo
                    checkingAccount.DisplayAccountInfo();
                    Console.WriteLine("\nFinish Operation (Press Enter)\n---------------------------------------");
                    Console.ReadKey();
                    break;
                case 4:  // AddInterest
                    Console.WriteLine("AddInterest Option, Not available =/");
                    Console.WriteLine("\nFinish Operation (Press Enter)\n---------------------------------------");
                    Console.ReadKey();
                    break;
                case 5:  // Exit Program

                    Console.WriteLine($"Final Status:\nAccount: {checkingAccount.AccountHolder}\nBalance: {checkingAccount.Balance}");
                    Console.ReadKey();

                    _switch = false;
                    break;
                default:
                    Console.WriteLine("Default");
                    break;
            }
        }

        // Aqui eu quero chamar a função de salvar no excel

    }
}
