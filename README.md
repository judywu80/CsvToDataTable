# CsvToDataTable
via List in winform <br>
※CsvToDataGridView主要是讀取csv檔，寫入(或採DataSource)到DataGridView，再儲存到文字檔，本作品為增加儲存到資料庫Emplo的資料表Employee。

- Form功能目標：
1. btn1為讀取csv文字檔，並顯示資料於DataGridView。
2. btn2將DataGridView的內容存於文字檔。
3. btn3將DataGridView的內容存於資料庫Emplo中的資料表Employee內。
4. btn4將資料表內容顯示於DataGridView。

![image](https://user-images.githubusercontent.com/122083665/235310409-98b65718-6035-4af2-a008-0c3a7e745f17.png)

- 資料讀取到DataGridView與排序作法參考：
1. 資料讀取到DataGridView方式(5.1既定功能)：
採用DataGridView.DataSource= List<T> 

2. 各欄位都可以排序，並且有升冪、降冪(5.1既定功能)：
DataSource=List<T>; 需使用List泛形排序(OrderBy、Delegate、Labmda…)

- [存資料表]作業方向參考：
1. 存入實體資料表方式：
寫入實體Table之前，採清除原資料表。
使用SQL語法直接寫到實體資料表
採用DataAdapter.Updata()；

2. 寫入資料表的資料來源：
DataGridView → Table
List<T> → Table

- [開Table檔]作業方向參考：(本功能僅作資料顯示，不作資料更新)
使用DataReader -> DataTable。
DataAdapter.Fill(DataTable); <br>
DataAdapter.Fill(DataSet , DataTable);  ->  DataTable=DataSet.DataTable[0] <br>
※顯示於DataGridView，建議使用DataSource。


# Csv To DataGridView
Form基本架構：原置放ListBox 作為載入文字檔用；文字檔轉換陣列後置放於DatagGridView內。

作品特色：<br>
- 一般級：
1. 於「FormLoad」事件(Event)檔，將專案目錄內的B01.csv檔載入到listBox1。
2. 藉由「執行轉換」鍵或其他事件、方法，將listBox1(A項完成的)內容，轉換成陣列並寫入DataGridView。
3. 分割文字可使用String.Split(',')。

- 實用級：
1. 設定DataGridView 使奇、偶數列不同背景色之功能(如下圖)。
![image](https://user-images.githubusercontent.com/122083665/235310165-2ff6004e-5d15-4399-8f69-60e812cd9deb.png)

- 進階級：
1. 利用「存文字檔」 Button事件，將DataGridView內容寫入文字檔test.txt於專題目錄內，並以ComBoBox元件，供使用者選擇該文字檔的分隔符號(逗號、分號、空白、TAB、冒號)，見如下二圖。
![image](https://user-images.githubusercontent.com/122083665/235310722-e31bd7e2-5e47-4c6b-87a2-4002100e060f.png)

2. 讓各欄位都可以排序，並且可以有升冪、降冪，DataGridView有預設排序功能，調整使其適用於本範例。

**外加功能(Option)**
- 本題改採用類別作法，由主程式提供檔案名稱(B01.csv)，後續的文字分割以及置入DataGridView工作(或採DataSource)，皆由該類別方法完成。

- 外加泛型自訂排序功能：（任選部份排序功能）
※建議使用List<T>，<T>型態自選，排序表達方式參考：
1. 自訂的 Class 繼承自 IComparable 以便具有 Sorting 功能 (List.Sort);
2. 定義 數個Class Compare 預設為各欄位作排序，如compareByName….
3. 使用委託 (Delegate) 排序
4. 使用 Lambda 表達式
5. 使用 List.OrderBy()，其中如果能夠使用由欄位名稱反射欄位屬性，再做排序，程式簡便性較佳。<br>
※以上排序方法，並非只選一種，可多種合用，如某欄位1使用排序表達方式1，某幾個欄位使用排序表達方式2…。
