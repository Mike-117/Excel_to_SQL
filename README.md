# Excel to SQL
 This is a console application that takes an existing Excel worksheet and translates it into a SQL Database.
 
 My initial desire was to use a single C# method to perform the operation, so I searched for an appropriate one. My research led me to the OLE DB Provider, and I tested various examples of it in action. But none of these worked for me, and with some added research I discovered that the method had been deprecated, and would no longer work.
 
 With time on the project running short, I decided not to look into any new methods but apply what I already knew. So I constructed two programs - a method to turn an Excel .xlsx file into a C# datatable, and another one to turn a C# datatable into a SQL database table. Then I simply combined them into one program, and I got the result I wanted.
