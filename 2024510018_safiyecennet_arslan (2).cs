using System;
using System.IO;
namespace MiniExcelProject
{
    class Program
    {
        //global variables
        static string[,] MaxSheet = new string[15, 10];
        static string[,] SheetTypes = new string[15, 10]; 

        static int CurrentRowCount = 8; //başlangıç satır
        static int CurrentColumnCount = 5; //başlangıç sütun
        static void InitializeSheet() 
        {
            for (int i = 0; i < 15; i++)
            {
                for (int j = 0; j < 10; j++)
                {
                    MaxSheet[i, j] = "";
                    SheetTypes[i, j] = "unassigned";
                }
            }
        }
        static void PrintSheet()
        {
            string Alfabe = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            Console.WriteLine();
            Console.Write("     ");
            for (int j = 0; j < CurrentColumnCount; j++)
            {
                Console.Write("    " + Alfabe[j] + "    ");
            }
            Console.WriteLine();
            Console.Write("    +");
            for (int j = 0; j < CurrentColumnCount; j++)
            {
                Console.Write("--------+");
            }
            Console.WriteLine(); 
            for (int i = 0; i < CurrentRowCount; i++)
            {
                Console.Write(i + 1); 
                if ((i + 1) < 10)
                    Console.Write("   "); 
                else
                    Console.Write("  ");  

                Console.Write("|");
                for (int j = 0; j < CurrentColumnCount; j++)
                {
                    string Veri = MaxSheet[i, j];
                    if (Veri == null)
                        Veri = ""; 
                    if (Veri.Length > 8) 
                    {
                        Veri = Veri.Substring(0, 8);
                    }
                    Console.Write(Veri);
                    int GerekenBosluk = 8 - Veri.Length;
                    for (int k = 0; k < GerekenBosluk; k++)
                    {
                        Console.Write(" ");
                    }
                    Console.Write("|");
                }
                Console.WriteLine(); 
                Console.Write("    +");
                for (int j = 0; j < CurrentColumnCount; j++)
                    Console.Write("--------+");
                Console.WriteLine();
            }
        }
        static void ProcessCommand(string input)
        {
            int ParantezYeri = -1;
            for (int i = 0; i < input.Length; i++)
            {
                if (input[i] == '(')
                {
                    ParantezYeri = i;
                    break;
                }
            }
            if (ParantezYeri == -1)
            {
                if (CheckIfCellAddress(input))
                {
                    ShowCellContent(input);
                    return;
                }
                Console.WriteLine("Hatalı komut formatı veya parantez eksik!");
                return;
            }
            string Operation = input.Substring(0, ParantezYeri);
            string ParametersPart = input.Substring(ParantezYeri + 1);
            if (ParametersPart.Length > 0 && ParametersPart[ParametersPart.Length - 1] == ')')
            {
                ParametersPart = ParametersPart.Substring(0, ParametersPart.Length - 1);
            }
            string[] Parameters = SplitArguments(ParametersPart);
            if (Operation == "AssignValue")
                AssignValue(Parameters);
            else if (Operation == "ClearCell")
                ClearCell(Parameters);
            else if (Operation == "ClearAll")
                ClearAll();
            else if (Operation == "save")
            {
                SaveToFile();
                Console.WriteLine("Dosya başarıyla kaydedildi!");
            }
            else if (Operation == "AddRow")
                AddRow(Parameters);
            else if (Operation == "AddColumn")
                AddColumn(Parameters);
            else if (Operation == "Copy")
                Copy(Parameters);
            else if (Operation == "CopyColumn")
                CopyColumn(Parameters);
            else if (Operation == "CopyRow")
                CopyRow(Parameters);
            else if (Operation == "X")
                XOperation(Parameters);
            else if (Operation == "XColumn")
                XColumn(Parameters);
            else if (Operation == "XRow")
                XRow(Parameters);
            else if (Operation == "*")
                MultiplicationOperation(Parameters);
            else if (Operation == "+")
                AdditionOperation(Parameters);
            else if (Operation == "/")
                DivisionOperation(Parameters);
            else if (Operation == "-")
                SubtractionOperation(Parameters);
            else if (Operation == "#")
                EncryptionOperation(Parameters);
            else if (CheckIfCellAddress(Operation))
                ShowCellContent(Operation);
            else
                Console.WriteLine("Bilinmeyen komut!");
        }
        static string[] SplitArguments(string ParametersText)
        {
            string[] parts = ParametersText.Split(',');
            for (int i = 0; i < parts.Length; i++)
            {
                string Word = parts[i];
                int start = 0; //baştaki boşlukları atmak için
                while (start < Word.Length && Word[start] == ' ')
                {
                    start++;
                }
                int final = Word.Length - 1; //sondaki boşlukları atmak için
                while (final >= 0 && Word[final] == ' ')
                {
                    final--;
                }
                string CleanWord = "";
                if (start <= final)
                {
                    for (int k = start; k <= final; k++)
                    {
                        CleanWord += Word[k];
                    }
                }
                parts[i] = CleanWord;
            }
            return parts;
        }
        static int GetRowIndex(string CellName) //alınan inputtan kaçıncı satır olduğunu tespit ederiz bu fonksiyonla
        {
            string RowPart = CellName.Substring(1);
            if (RowPart == "")
                return -1;
            int sonucSayi = 0;
            for (int i = 0; i < RowPart.Length; i++)
            {
                char c = RowPart[i];
                if (c < '0' || c > '9')
                {
                    return -1; 
                }
                int rakamDegeri = c - '0';
                sonucSayi = (sonucSayi * 10) + rakamDegeri; //örneğin verilen satır numarası 2 basamaklıysa bu yöntemle basitçe okuyabiliriz
            }
            return sonucSayi - 1; //indeksin satır numarasından 1 eksik olmasından dolayı kaynaklanır
        }
        static int GetColIndex(string CellName)
        {
            char ColChar = CellName.ToUpper()[0];
            return ColChar - 'A';
        }
        static bool CheckIfCellAddress(string input)
        {
            if (input.Length < 2)
                return false;
            char c = input.ToUpper()[0];
            return (c >= 'A' && c <= 'J');
        }
        static bool ValidPos(int r, int c) //sınırlar içinde mi kontrol eder
        {
            return r >= 0 && r < CurrentRowCount && c >= 0 && c < CurrentColumnCount;
        }
        static void ShowCellContent(string CellName)
        {
            int r = GetRowIndex(CellName);
            int c = GetColIndex(CellName);
            if (ValidPos(r, c))
            {
                Console.WriteLine(MaxSheet[r, c]);
            }
            else
            {
                Console.WriteLine("Out of bounds!");
            }
        }
        static void AssignValue(string[] Parameters)
        {
            if (Parameters.Length != 3)
            {
                Console.WriteLine("Invalid parameter count!");
                return;
            }
            string Cell = Parameters[0];
            string Type = Parameters[1].ToLower();
            string Val = Parameters[2];
            int r = GetRowIndex(Cell);
            int c = GetColIndex(Cell);
            if (!ValidPos(r, c))
            {
                Console.WriteLine("Out of bounds exception!");
                return;
            }
            MaxSheet[r, c] = Val;
            SheetTypes[r, c] = Type;
            Console.WriteLine("Operation is done!");
        }
        static void ClearCell(string[] Parameters)
        {
            string Cell = Parameters[0];
            int r = GetRowIndex(Cell);
            int c = GetColIndex(Cell);
            if (ValidPos(r, c))
            {
                MaxSheet[r, c] = "";
                SheetTypes[r, c] = "unassigned";
                Console.WriteLine("Operation is done!");
            }
            else
                Console.WriteLine("Out of bounds!");
        }
        static void ClearAll()
        {
            InitializeSheet();
        }
        static void AddRow(string[] Parameters)
        {
            //Tablo tamamen dolu mu kontrol eder
            if (CurrentRowCount >= 15)
            {
                Console.WriteLine("Max row size reached!");
                return;
            }
            string RowString = Parameters[0];
            int RowNo = 0;
            for (int i = 0; i < RowString.Length; i++)
            {
                RowNo = RowNo * 10 + (RowString[i] - '0');
            }
            int TargetIndex = RowNo - 1; 
            string direction = Parameters[1]; //up veya down
            int InsertIndex = 0;
            if (direction == "down")
            {
                InsertIndex = TargetIndex + 1; 
            }
            else
            {
                InsertIndex = TargetIndex; 
            }
            if (InsertIndex < 0 || InsertIndex > CurrentRowCount)
            {
                Console.WriteLine("Illegal position assignment!");
                return;
            }
            for (int i = CurrentRowCount; i > InsertIndex; i--)
            {
                for (int j = 0; j < CurrentColumnCount; j++)
                {
                    MaxSheet[i, j] = MaxSheet[i - 1, j];
                    SheetTypes[i, j] = SheetTypes[i - 1, j];
                }
            }
            for (int j = 0; j < CurrentColumnCount; j++)
            {
                MaxSheet[InsertIndex, j] = "";
                SheetTypes[InsertIndex, j] = "unassigned";
            }
            CurrentRowCount++;
            Console.WriteLine("Operation is done!");
        }
        static void AddColumn(string[] Parameters)
        {
            if (CurrentColumnCount >= 10)
            {
                Console.WriteLine("Max column size reached!");
                return;
            }
            char ColChar = Parameters[0][0];
            int ColNo = ColChar - 'A';
            string Direction = Parameters[1].ToLower();
            int InsertIndex = 0;
            if (Direction == "right")
            {
                InsertIndex = ColNo + 1;
            }
            else
            {
                InsertIndex = ColNo;
            }
            if (InsertIndex < 0 || InsertIndex > CurrentColumnCount)
            {
                Console.WriteLine("Illegal position assignment!");
                return;
            }
            for (int j = CurrentColumnCount; j > InsertIndex; j--)
            {

                for (int i = 0; i < CurrentRowCount; i++)
                {
                    MaxSheet[i, j] = MaxSheet[i, j - 1];
                    SheetTypes[i, j] = SheetTypes[i, j - 1];
                }
            }
            for (int i = 0; i < CurrentRowCount; i++)
            {
                MaxSheet[i, InsertIndex] = "";
                SheetTypes[i, InsertIndex] = "unassigned";
            }
            CurrentColumnCount++;
            Console.WriteLine("Operation is done!");
        }
        static void Copy(string[] Parameters)
        {
            int r1 = GetRowIndex(Parameters[0]);
            int c1 = GetColIndex(Parameters[0]);
            int r2 = GetRowIndex(Parameters[1]);
            int c2 = GetColIndex(Parameters[1]);
            if (ValidPos(r1, c1) && ValidPos(r2, c2))
            {
                MaxSheet[r2, c2] = MaxSheet[r1, c1];
                SheetTypes[r2, c2] = SheetTypes[r1, c1];
                Console.WriteLine("Operation is done!");
            }
            else
                Console.WriteLine("Illegal position assignment!");
        }
        static void CopyColumn(string[] Parameters)
        {
            int c1 = Parameters[0][0] - 'A';
            int c2 = Parameters[1][0] - 'A';
            if (c1 >= 0 && c1 < CurrentColumnCount && c2 >= 0 && c2 < CurrentColumnCount)
            {
                for (int i = 0; i < CurrentRowCount; i++)
                {
                    MaxSheet[i, c2] = MaxSheet[i, c1];
                    SheetTypes[i, c2] = SheetTypes[i, c1];
                }
                Console.WriteLine("Operation is done!");
            }
            else
                Console.WriteLine("Illegal position assignment!");
        }
        static void CopyRow(string[] Parameters)
        {
            string RowStr1 = Parameters[0]; 
            int r1 = 0;
            for (int k = 0; k < RowStr1.Length; k++)
            {
                r1 = r1 * 10 + (RowStr1[k] - '0');
            }
            r1 = r1 - 1;
            string RowStr2 = Parameters[1];
            int r2 = 0;
            for (int k = 0; k < RowStr2.Length; k++)
            {
                r2 = r2 * 10 + (RowStr2[k] - '0');
            }
            r2 = r2 - 1;
            if (r1 >= 0 && r1 < CurrentRowCount && r2 >= 0 && r2 < CurrentRowCount)
            {
                for (int j = 0; j < CurrentColumnCount; j++)
                {
                    MaxSheet[r2, j] = MaxSheet[r1, j];
                    SheetTypes[r2, j] = SheetTypes[r1, j];
                }
                Console.WriteLine("Operation is done!");
            }
            else
            {
                Console.WriteLine("Illegal position assignment!");
            }
        }
        static void XOperation(string[] Parameters)
        {
            Copy(Parameters);
            string[] ClearParameters = { Parameters[0] };
            ClearCell(ClearParameters);
        }
        static void XColumn(string[] Parameters)
        {
            CopyColumn(Parameters);
            int c1 = Parameters[0][0] - 'A';
            for (int i = 0; i < CurrentRowCount; i++)
            {
                MaxSheet[i, c1] = "";
                SheetTypes[i, c1] = "unassigned";
            }
        }
        static void XRow(string[] Parameters)
        {
            CopyRow(Parameters);
            string RowStr = Parameters[0];
            int r1 = 0;
            for (int k = 0; k < RowStr.Length; k++)
            {
                r1 = r1 * 10 + (RowStr[k] - '0');
            }
            r1 = r1 - 1;
            for (int j = 0; j < CurrentColumnCount; j++)
            {
                MaxSheet[r1, j] = "";
                SheetTypes[r1, j] = "unassigned";
            }
        }
        static void MultiplicationOperation(string[] Parameters)
        {
            int r1 = GetRowIndex(Parameters[0]);
            int c1 = GetColIndex(Parameters[0]);
            int r2 = GetRowIndex(Parameters[1]);
            int c2 = GetColIndex(Parameters[1]);
            int ro = GetRowIndex(Parameters[2]);
            int co = GetColIndex(Parameters[2]);
            if (!ValidPos(r1, c1) || !ValidPos(r2, c2) || !ValidPos(ro, co))
            {
                Console.WriteLine("Out of bounds!");
                return;
            }
            string t1 = SheetTypes[r1, c1];
            string t2 = SheetTypes[r2, c2];
            string v1 = MaxSheet[r1, c1];
            string v2 = MaxSheet[r2, c2];
            if (t1 == "unassigned" || t2 == "unassigned")
            {
                Console.WriteLine("Illegal operation! Unassigned cell.");
                return;
            }
            if (t1 == "integer" && t2 == "integer")
            {
                int Num1 = 0;
                int IsNeg1 = 1;
                int Start1 = 0;
                if (v1[0] == '-')
                {
                    IsNeg1 = -1;
                    Start1 = 1;
                }
                for (int i = Start1; i < v1.Length; i++)
                {
                    Num1 = Num1 * 10 + (v1[i] - '0');
                }
                Num1 = Num1 * IsNeg1;
                int Num2 = 0;
                int IsNeg2 = 1;
                int Start2 = 0;
                if (v2[0] == '-')
                {
                    IsNeg2 = -1;
                    Start2 = 1;
                }
                for (int i = Start2; i < v2.Length; i++)
                {
                    Num2 = Num2 * 10 + (v2[i] - '0');
                }
                Num2 = Num2 * IsNeg2;
                int multi = Num1 * Num2;
                MaxSheet[ro, co] = "" + multi;
                SheetTypes[ro, co] = "integer";
                Console.WriteLine("Operation is done!");
            }
            else if (t1 == "string" && t2 == "string")
            {
                Console.WriteLine("Illegal operation! String String operation is not allowed!");
            }
            else
            {
                string textToRepeat = "";
                string numberString = "";
                if (t1 == "string")
                {
                    textToRepeat = v1;
                    numberString = v2;
                }
                else
                {
                    textToRepeat = v2;
                    numberString = v1;
                }
                int k = 0;
                int IsNegK = 1;
                int StartK = 0;
                if (numberString[0] == '-')
                {
                    IsNegK = -1;
                    StartK = 1;
                }
                for (int i = StartK; i < numberString.Length; i++)
                {
                    k = k * 10 + (numberString[i] - '0');
                }
                k = k * IsNegK;
                if (k < 0)
                {
                    k = k * -1;
                    string reversedText = "";
                    for (int i = textToRepeat.Length - 1; i >= 0; i--)
                    {
                        reversedText += textToRepeat[i];
                    }
                    textToRepeat = reversedText;
                }
                string result = "";
                for (int i = 0; i < k; i++)
                {
                    result += textToRepeat;
                }
                MaxSheet[ro, co] = result;
                SheetTypes[ro, co] = "string";
                Console.WriteLine("Operation is done! Output: " + result);
            }
        }
        static void AdditionOperation(string[] Parameters)
        {
            int OutputIdx = Parameters.Length - 1;
            int ro = GetRowIndex(Parameters[OutputIdx]);
            int co = GetColIndex(Parameters[OutputIdx]);
            if (!ValidPos(ro, co))
            {
                Console.WriteLine("Output out of bounds!");
                return;
            }
            bool AnyString = false;
            for (int i = 0; i < OutputIdx; i++)
            {
                int r = GetRowIndex(Parameters[i]);
                int c = GetColIndex(Parameters[i]);
                if (!ValidPos(r, c) || SheetTypes[r, c] == "unassigned")
                {
                    Console.WriteLine("Illegal or Unassigned cell!");
                    return;
                }
                if (SheetTypes[r, c] == "string")
                {
                    AnyString = true;
                }
            }
            if (AnyString)
            {
                Console.Write("Concatenation operation will be applied!\nPlease select letter case (up/low): ");
                string LetterCase = Console.ReadLine(); 
                string rawResult = "";
                for (int i = 0; i < OutputIdx; i++)
                {
                    int r = GetRowIndex(Parameters[i]);
                    int c = GetColIndex(Parameters[i]);
                    rawResult += MaxSheet[r, c];
                }
                string finalResult = "";
                if (LetterCase == "up")
                {
                    for (int k = 0; k < rawResult.Length; k++)
                    {
                        char harf = rawResult[k];
                        if (harf >= 'a' && harf <= 'z')
                        {
                            finalResult += (char)(harf - 32);
                        }
                        else
                        {
                            finalResult += harf; 
                        }
                    }
                }
                else if (LetterCase == "low")
                {
                    for (int k = 0; k < rawResult.Length; k++)
                    {
                        char harf = rawResult[k];
                        if (harf >= 'A' && harf <= 'Z')
                        {
                            finalResult += (char)(harf + 32);
                        }
                        else
                        {
                            finalResult += harf;
                        }
                    }
                }
                else
                {
                    finalResult = rawResult; 
                }
                MaxSheet[ro, co] = finalResult;
                SheetTypes[ro, co] = "string";
            }
            else
            {
                int sum = 0;
                for (int i = 0; i < OutputIdx; i++)
                {
                    int r = GetRowIndex(Parameters[i]);
                    int c = GetColIndex(Parameters[i]);
                    string val = MaxSheet[r, c];

                    int currentNum = 0;
                    int isNeg = 1;
                    int startIdx = 0;
                    if (val[0] == '-')
                    {
                        isNeg = -1;
                        startIdx = 1;
                    }
                    for (int k = startIdx; k < val.Length; k++)
                    {
                        currentNum = currentNum * 10 + (val[k] - '0');
                    }
                    currentNum = currentNum * isNeg; 
                    sum += currentNum;
                }
                MaxSheet[ro, co] = "" + sum;
                SheetTypes[ro, co] = "integer";
            }
            Console.WriteLine("Operation is done!");
        }
        static void DivisionOperation(string[] Parameters)
        {
            int r1 = GetRowIndex(Parameters[0]);
            int c1 = GetColIndex(Parameters[0]);
            int r2 = GetRowIndex(Parameters[1]);
            int c2 = GetColIndex(Parameters[1]);
            int ro = GetRowIndex(Parameters[2]); 
            int co = GetColIndex(Parameters[2]);
            if (!ValidPos(r1, c1) || !ValidPos(r2, c2) || !ValidPos(ro, co))
            {
                Console.WriteLine("Out of bounds!");
                return;
            }
            string t1 = SheetTypes[r1, c1];
            string t2 = SheetTypes[r2, c2];
            string v1 = MaxSheet[r1, c1];
            string v2 = MaxSheet[r2, c2];
            if (t1 == "integer" && t2 == "integer")
            {
                int Num1 = 0;
                int IsNeg1 = 1;
                int Start1 = 0;
                if (v1[0] == '-')
                {
                    IsNeg1 = -1;
                    Start1 = 1;
                }
                for (int i = Start1; i < v1.Length; i++)
                {
                    Num1 = Num1 * 10 + (v1[i] - '0');
                }
                Num1 = Num1 * IsNeg1;
                int Num2 = 0;
                int IsNeg2 = 1;
                int Start2 = 0;
                if (v2[0] == '-')
                {
                    IsNeg2 = -1;
                    Start2 = 1;
                }
                for (int i = Start2; i < v2.Length; i++)
                {
                    Num2 = Num2 * 10 + (v2[i] - '0');
                }
                Num2 = Num2 * IsNeg2;
                if (Num2 == 0)
                {
                    Console.WriteLine("Divide by zero error!");
                    return;
                }
                int res = Num1 / Num2;
                MaxSheet[ro, co] = res.ToString();
                SheetTypes[ro, co] = "integer";
                Console.WriteLine("Operation is done!");
            }
            else if (t1 == "string" && t2 == "string")
            {
                Console.WriteLine("Illegal operation! String String not allowed.");
            }
            else
            {
                string text = "";
                string NumberStr = "";
                if (t1 == "string")
                {
                    text = v1;
                    NumberStr = v2;
                }
                else
                {
                    text = v2;
                    NumberStr = v1;
                }
                int k = 0;
                int IsNegK = 1;
                int StartK = 0;
                if (NumberStr[0] == '-') 
                { 
                    IsNegK = -1; 
                    StartK = 1;
                }
                for (int i = StartK; i < NumberStr.Length; i++)
                {
                    k = k * 10 + (NumberStr[i] - '0');
                }
                k = k * IsNegK; 
                int AbsK = Math.Abs(k);
                if (AbsK == 0) 
                    AbsK = 1;
                int PartSize = text.Length / AbsK;
                string Result = "";
                if (k > 0)
                {
                    Result = text.Substring(0, PartSize);
                }
                else
                {
                    Result = text.Substring(text.Length - PartSize);
                }
                MaxSheet[ro, co] = Result;
                SheetTypes[ro, co] = "string";
                Console.WriteLine("Operation is done! Result: " + Result);
            }
        }
        static void SubtractionOperation(string[] Parameters)
        {
            int ro = GetRowIndex(Parameters[2]);
            int co = GetColIndex(Parameters[2]);
            if (!ValidPos(ro, co))
            {
                Console.WriteLine("Output out of bounds!");
                return;
            }
            string Parameter1 = Parameters[0];
            string v1 = "";
            string t1 = "";
            int rTest1 = GetRowIndex(Parameter1);
            int cTest1 = GetColIndex(Parameter1);
            if (ValidPos(rTest1, cTest1))
            {
                v1 = MaxSheet[rTest1, cTest1];
                t1 = SheetTypes[rTest1, cTest1];
            }
            else
            {
                v1 = Parameter1;
                bool IsNumber = true;
                int StartCheck = 0;
                if (v1.Length > 0 && v1[0] == '-') 
                    StartCheck = 1; 
                for (int i = StartCheck; i < v1.Length; i++)
                {
                    if (v1[i] < '0' || v1[i] > '9')
                    {
                        IsNumber = false;
                    }
                }
                if (IsNumber && v1 != "")
                    t1 = "integer";
                else t1 = "string";
            }
            string Parameter2 = Parameters[1];
            string v2 = "";
            string t2 = "";
            int rTest2 = GetRowIndex(Parameter2);
            int cTest2 = GetColIndex(Parameter2);
            if (ValidPos(rTest2, cTest2))
            {
                v2 = MaxSheet[rTest2, cTest2];
                t2 = SheetTypes[rTest2, cTest2];
            }
            else
            {
                v2 = Parameter2;
                bool IsNumber = true;
                int StartCheck = 0;
                if (v2.Length > 0 && v2[0] == '-') StartCheck = 1;
                for (int i = StartCheck; i < v2.Length; i++)
                {
                    if (v2[i] < '0' || v2[i] > '9') IsNumber = false;
                }
                if (IsNumber && v2 != "") t2 = "integer";
                else t2 = "string";
            }
            if (t1 == "integer" && t2 == "integer")
            {
                int Num1 = 0;
                int IsNeg1 = 1;
                int s1 = 0;
                if (v1[0] == '-') { IsNeg1 = -1; s1 = 1; }
                for (int i = s1; i < v1.Length; i++) Num1 = Num1 * 10 + (v1[i] - '0');
                Num1 = Num1 * IsNeg1;
                int Num2 = 0;
                int IsNeg2 = 1;
                int s2 = 0;
                if (v2[0] == '-')
                {
                    IsNeg2 = -1;
                    s2 = 1;
                }
                for (int i = s2; i < v2.Length; i++)
                    Num2 = Num2 * 10 + (v2[i] - '0');
                Num2 = Num2 * IsNeg2;
                int res = Num1 - Num2;
                MaxSheet[ro, co] = res.ToString();
                SheetTypes[ro, co] = "integer";
            }
            else if (t1 == "string" && t2 == "string")
            {
                string Longer = "";
                string Shorter = "";
                if (v1.Length >= v2.Length)
                {
                    Longer = v1;
                    Shorter = v2;
                }
                else
                {
                    Longer = v2;
                    Shorter = v1;
                }
                string result = Longer.Replace(Shorter, "");
                MaxSheet[ro, co] = result;
                SheetTypes[ro, co] = "string";
            }
            else
            {
                string text = "";
                string AsciiStr = "";
                if (t1 == "string")
                {
                    text = v1;
                    AsciiStr = v2;
                }
                else
                {
                    text = v2;
                    AsciiStr = v1;
                }
                int AsciiCode = 0;
                for (int i = 0; i < AsciiStr.Length; i++)
                {
                    AsciiCode = AsciiCode * 10 + (AsciiStr[i] - '0');
                }
                if (AsciiCode < 33 || AsciiCode > 126)
                {
                    Console.WriteLine("Illegal operation! Integer must be in [33, 126]");
                    return;
                }
                char Target = (char)AsciiCode;
                string Result = "";
                for (int i = 0; i < text.Length; i++)
                {
                    char ch = text[i];
                    if (ch != Target)
                    {
                        Result += ch;
                    }
                }
                MaxSheet[ro, co] = Result;
                SheetTypes[ro, co] = "string";
            }
            Console.WriteLine("Operation is done!");
        }
        static void EncryptionOperation(string[] Parameters)
        {
            int r1 = GetRowIndex(Parameters[0]);
            int c1 = GetColIndex(Parameters[0]);
            int r2 = GetRowIndex(Parameters[1]);
            int c2 = GetColIndex(Parameters[1]);
            int ro = GetRowIndex(Parameters[2]); 
            int co = GetColIndex(Parameters[2]);
            if (!ValidPos(r1, c1) || !ValidPos(r2, c2) || !ValidPos(ro, co))
            {
                return; 
            }
            string t1 = SheetTypes[r1, c1];
            string t2 = SheetTypes[r2, c2];
            string v1 = MaxSheet[r1, c1];
            string v2 = MaxSheet[r2, c2];
            if ((t1 == "string" && t2 == "string") || (t1 == "integer" && t2 == "integer"))
            {
                Console.WriteLine("Illegal operation! Need one string and one integer.");
                return;
            }
            string text = "";
            string ShiftString = "";
            if (t1 == "string")
            {
                text = v1;          
                ShiftString = v2;   
            }
            else
            {
                text = v2;
                ShiftString = v1;
            }
            int Shift = 0;
            int IsNeg = 1;      
            int StartIdx = 0;   
            if (ShiftString[0] == '-')
            {
                IsNeg = -1;
                StartIdx = 1; 
            }
            for (int i = StartIdx; i < ShiftString.Length; i++)
            {
                Shift = Shift * 10 + (ShiftString[i] - '0');
            }
            Shift = Shift * IsNeg; 
            if (Shift < -20 || Shift > 30)
            {
                Console.WriteLine("Illegal operation! Shift must be in [-20, 30]");
                return;
            }
            string Result = "";
            for (int i = 0; i < text.Length; i++)
            {
                char OriginalChar = text[i];
                char EncryptedChar = (char)(OriginalChar + Shift);
                Result += EncryptedChar;
            }
            MaxSheet[ro, co] = Result;
            SheetTypes[ro, co] = "string";
            Console.WriteLine("Operation is done! Encrypted: " + Result);
        }
        static void SaveToFile()
        {
            StreamWriter sw = new StreamWriter("spreadsheet.txt");
            for (int j = 0; j < CurrentColumnCount; j++)
            {
                sw.Write(((char)('A' + j)) + "\t");
            }
            sw.WriteLine(); 
            for (int i = 0; i < CurrentRowCount; i++)
            {
                sw.Write((i + 1) + "\t"); 
                for (int j = 0; j < CurrentColumnCount; j++)
                {
                    sw.Write(MaxSheet[i, j] + "\t"); 
                }
                sw.WriteLine(); 
            }
            sw.Close();
        }
        static void Main(string[] args)
        {
            InitializeSheet();
            string command = "";
            while (true)
            {
                PrintSheet();
                Console.WriteLine("\nEnter command (or 'exit' to close):");
                Console.Write(">> ");
                command = Console.ReadLine();
                if (command.ToLower() == "exit")
                {
                    Console.WriteLine("Saving files...");
                    SaveToFile();
                    break;
                }
                ProcessCommand(command);
            }
        }
    }
}

