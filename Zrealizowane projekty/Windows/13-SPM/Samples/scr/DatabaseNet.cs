using System;
using System.Data;
using System.Collections.Generic;
using System.IO;
using MyLib.MyBase;
using MyLib.Crc32;
using SQLDatabase.Net.SQLDatabaseClient;
using OlimpComponents;
using System.Reflection;
using MyLib.Security;
using Newtonsoft.Json;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;

namespace Aplikacja
{
   public class DatabaseNet : MyClass
   {
      protected SqlDatabaseConnection _Connection;
      protected string _PhraseKey;

      public string DirDb { get; private set; }
      public string DatabaseName { get; private set; }
      public DateTime ExpireDate { get; private set; }
      public string PasswordHash { get; private set; }
      public string KeyXsec { get; private set; }

      public List<string> Tables { get; private set; } = new List<string>();
      //
      public DatabaseNet(string dirDb, string phraseKey)
      {
         DirDb = dirDb;
         _PhraseKey = phraseKey;
      }

      ~DatabaseNet()
      {
         Close();   
      }

      public SqlDatabaseConnection Connection { get { return _Connection; } }

      public bool Open(string NameDb, String Password)
      {
         bool result = false;
         try
         {
               if (_Connection != null) Close();

               string fileDb = Path.ChangeExtension(Path.Combine(DirDb, NameDb), ".db");
               string fileKey = Path.ChangeExtension(Path.Combine(DirDb, NameDb), ".key");
               if (!File.Exists(fileDb)) throw new InvalidOperationException($"Brak pliku bazy danych: {fileDb}");
               if (!File.Exists(fileKey)) throw new InvalidOperationException($"Brak pliku klucza: {fileKey}");

               string bodyKey = File.ReadAllText(fileKey);
               string bodyKeyDecrypt = XTEA.Decrypt(bodyKey, Const.KEY_SecurePhraseFile);
               SignatureFile signature = JsonConvert.DeserializeObject<SignatureFile>(bodyKeyDecrypt) as SignatureFile;

            if (signature.HashPassword == HashPassword(Password))
            {
               string mixedKey = KeyBuilder(Password);
               try
               {
                  var connection = new SqlDatabaseConnection();
                  connection.ConnectionString = $"SchemaName=db;uri=file://{fileDb}";
                  connection.Open();
                  if (connection.State == ConnectionState.Open)
                  {
                     using (SqlDatabaseCommand command = new SqlDatabaseCommand())
                     {
                        command.Connection = connection;
                        command.CommandText = $"SYSCMD Key = '{mixedKey}'";
                        command.ExecuteScalar();
                        // pobierz listę tabel
                        Tables.Clear();
                        command.CommandText = $"SELECT tablename FROM SYS_OBJECTS Where type = 'table'";
                        using (var reader = command.ExecuteReader())
                        {
                           while (reader.Read()) Tables.Add(reader.GetString(0));
                        }
                        // udało się
                        _Connection = connection;
                        DatabaseName = NameDb;
                        ExpireDate = signature.ExpireDate;
                        PasswordHash = signature.HashPassword;
                        KeyXsec = signature.KeyXsec;
                        result = true;
                     }
                  }
                  else
                     connection?.Close();
               }
               catch
               {
                  throw;
               }
            }
         }
         catch (Exception ex)
         {
            base.RaiseErrorEvent(ex);
         }
         return result;
      }

      public void Close()
      {
         bool state = ((_Connection != null) && (_Connection.State == ConnectionState.Open));
         if (state) 
         {
            _Connection.Close();
            _Connection = null;
            Tables.Clear();
            DatabaseName = null;
            ExpireDate = default(DateTime);
            PasswordHash = null;
            KeyXsec = null;
         }
      }

      protected SqlDatabaseTransaction BeginTransaction(SqlDatabaseConnection connection)
      {
         bool state = ((connection != null) && (connection.State == ConnectionState.Open));
         return state ? connection.BeginTransaction() : null;
      }

      public SqlDatabaseTransaction BeginTransaction()
      {
         return BeginTransaction(_Connection);
      }

      public bool Commit(SqlDatabaseTransaction transaction)
      {
         bool result = false;
         bool state = ((_Connection != null) && (_Connection.State == ConnectionState.Open) && (transaction != null));
         if (state)
         {
            transaction.Commit();
            result = true;
         }
         return result;
      }

      public bool Rollback(SqlDatabaseTransaction transaction)
      {
         bool result = false;
         bool state = ((_Connection != null) && (_Connection.State == ConnectionState.Open) && (transaction != null));
         if (state)
         {
            transaction.Rollback();
            result = true;
         }
         return result;
      }

      private string PrepareSQL(string sqlSelect, NetParameters sqlParameters)
      {
         string sql = sqlSelect;
         
         List<string> blocks = new List<string>();
         const string tagFirst = "<param>";
         const string tagLast = "</param>";

         int ps = 0;
         int p1 = sql.IndexOf(tagFirst, ps, StringComparison.OrdinalIgnoreCase);
         int p2 = sql.IndexOf(tagLast, ps, StringComparison.OrdinalIgnoreCase);
         while ((p2 > p1) && (p1 > -1))
         {
            ps = p2 + tagLast.Length;
            blocks.Add(sql.Substring(p1, ps - p1));
            p1 = sql.IndexOf(tagFirst, ps, StringComparison.OrdinalIgnoreCase);
            p2 = sql.IndexOf(tagLast, ps, StringComparison.OrdinalIgnoreCase);
         }

         foreach (string block in blocks)
         {
            bool found = false;
            if (sqlParameters != null)
            {
               foreach (var par in sqlParameters.Items)
                  if (block.Contains(par.ParName))
                  {
                     found = true;
                     break;
                  }
            }
            if (!found)
            {
               sql = sql.Replace(block, String.Empty);
            }
         }
         sql = sql.Replace(tagFirst, String.Empty);
         sql = sql.Replace(tagLast, String.Empty);

         // dodatkowa zamiana listy int
         for (int iii = (sqlParameters?.Count ?? 0)-1; iii >= 0; iii--)
         {
            if (sqlParameters.Items[iii].ParValue is int[] intArray)
            {
               string result = string.Join(",", intArray);

               // Użycie Regex.Escape dla ochrony przed znakami specjalnymi
               string pattern = Regex.Escape(sqlParameters.Items[iii].ParName) + @"(?=[\s\)$]|$)";
               sql = Regex.Replace(sql, pattern, result);

               // Usuwanie przetworzonego elementu
               sqlParameters.Items.RemoveAt(iii);
            }
         }
         return sql;
      }

      protected DataTable Select(SqlDatabaseConnection connection, string sqlSelect, NetParameters sqlParameters = null) 
      {
         DataTable dt = new DataTable();
         try
         {
            bool state = ((connection != null) && (connection.State == ConnectionState.Open));
            if (!state) throw new InvalidOperationException("Brak połączenia do bazy danych SQL.");
            if (sqlSelect.IsNullOrEmpty()) throw new InvalidOperationException("Brak zapytania SQL.");

            string sql = PrepareSQL(sqlSelect, sqlParameters);

            using (SqlDatabaseCommand command = new SqlDatabaseCommand())
            {
               command.Connection = connection;
               command.CommandText = sql;
               if (sqlParameters != null)
               {
                  foreach (var par in sqlParameters.Items)
                     command.Parameters.AddWithValue(par.ParName, par.ParValue);
               }
               using (SqlDatabaseDataAdapter adapter = new SqlDatabaseDataAdapter())
               {
                  adapter.SelectCommand = command;
                  adapter.Fill(dt);
               }
            }
         }
         catch (Exception ex)
         {
            base.RaiseErrorEvent(ex);
         }
         return dt;
      }

      public DataTable Select(string sqlSelect, NetParameters sqlParameters = null)
      {
         return Select(_Connection, sqlSelect, sqlParameters);
      }

      protected List<T> Select<T>(SqlDatabaseConnection connection, string sqlSelect, NetParameters sqlParameters = null) where T : new()
      {
         List<T> result = new List<T>();
         try
         {
            bool state = ((connection != null) && (connection.State == ConnectionState.Open));
            if (!state) throw new InvalidOperationException("Brak połączenia do bazy danych SQL.");
            if (sqlSelect.IsNullOrEmpty()) throw new InvalidOperationException("Brak zapytania SQL.");
            string sql = PrepareSQL(sqlSelect, sqlParameters);

            using (SqlDatabaseCommand command = new SqlDatabaseCommand())
            {
               command.Connection = connection;
               command.CommandText = sql;
               if (sqlParameters != null)
               {
                  foreach (var par in sqlParameters.Items)
                     command.Parameters.AddWithValue(par.ParName, par.ParValue);
               }
               using (var reader = command.ExecuteReader())
               {
                  while (reader.Read())  // Tylko jeden rekord
                  {
                     T obj = new T();
                     foreach (var property in typeof(T).GetProperties())
                     {
                        var columnName = property.Name;
                        if (columnName != null)
                        {
                           try
                           {
                              var value = reader[columnName];
                              if (value != DBNull.Value)
                                 property.SetValue(obj, value);
                           }
                           catch { }
                        }
                     }
                     result.Add(obj);
                  }
               }

            }
         }
         catch (Exception ex)
         {
            base.RaiseErrorEvent(ex);
         }
         return result;
      }

      public List<T> Select<T>(string sqlSelect, NetParameters sqlParameters = null) where T : new()
      {
         return Select<T>(_Connection, sqlSelect, sqlParameters);
      }

      protected T SelectScalar<T>(SqlDatabaseConnection connection, string sqlSelect, NetParameters sqlParameters = null)
      {
         T result = default(T);
         try
         {
            bool state = ((connection != null) && (connection.State == ConnectionState.Open));
            if (!state) throw new InvalidOperationException("Brak połączenia do bazy danych SQL.");
            if (sqlSelect.IsNullOrEmpty()) throw new InvalidOperationException("Brak zapytania SQL.");

            string sql = PrepareSQL(sqlSelect, sqlParameters);

            using (SqlDatabaseCommand command = new SqlDatabaseCommand())
            {
               command.Connection = connection;
               command.CommandText = sql;
               if (sqlParameters != null)
               {
                  foreach (var par in sqlParameters.Items)
                     command.Parameters.AddWithValue(par.ParName, par.ParValue);
               }
               var readData = command.ExecuteScalar();
               result = (T)Convert.ChangeType(readData, typeof(T));
            }
         }
         catch (Exception ex)
         {
            base.RaiseErrorEvent(ex);
         }
         return result;
      }
      public T SelectScalar<T>(string sqlSelect, NetParameters sqlParameters = null)
      {
         return SelectScalar<T>(_Connection, sqlSelect, sqlParameters);
      }

      public static void SetPropertyValueByColumnName(object obj, string propertyName, object value)
      {
         var property = obj.GetType().GetProperty(propertyName, BindingFlags.Public | BindingFlags.Instance);
         if (property != null && property.CanWrite)
         {
            // Sprawdzenie, czy typ właściwości jest nullable
            var propertyType = Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType;
            // Sprawdzenie czy wartość to DBNull
            if (value == DBNull.Value) value = null;
            // Obsługuje nullable, konwertując wartość na odpowiedni typ
            object convertedValue = value == null ? null : Convert.ChangeType(value, propertyType);
            // Ustawienie wartości właściwości
            property.SetValue(obj, convertedValue);
         }
      }

      public T FillFromDataRow<T>(DataRow row) where T : new()
      {
         T result = new T();
         try
         {
            foreach (var property in typeof(T).GetProperties())
            {
               var columnName = property.Name;
               if (columnName != null)
               {
                  object value = null;
                  if (row != null && row.Table?.Columns.Contains(columnName) == true)
                  {
                     value = row[columnName];
                  }
                  else
                  {
                     // Typ właściwości (obsługa nullable i typów prostych)
                     var propertyType = Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType;
                     
                     // Ustawienie domyślnej wartości dla typów prostych
                     value = Nullable.GetUnderlyingType(property.PropertyType)== null && propertyType.IsValueType ? Activator.CreateInstance(propertyType) : null;
                  }
                  // Ustawienie wartości właściwości
                  SetPropertyValueByColumnName(result, property.Name, value);
               }
            }
         }
         catch (Exception ex)
         {
            base.RaiseErrorEvent(ex);
         }
         return result;
      }

      public List<T> FillFromDataTable<T>(DataTable dt) where T : new()
      {
         List<T> result = new List<T>();
         try
         {
            if (dt != null)
            {
               foreach (DataRow row in dt.Rows)  // Tylko jeden rekord
               {
                  T obj = new T();
                  foreach (var property in typeof(T).GetProperties())
                  {
                     var columnName = property.Name;
                     if (columnName != null)
                     {
                        object value = null;
                        if (row != null && row.Table?.Columns.Contains(columnName) == true)
                        {
                           value = row[columnName];
                        }
                        else
                        {
                           // Typ właściwości (obsługa nullable i typów prostych)
                           var propertyType = Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType;

                           // Ustawienie domyślnej wartości dla typów prostych
                           value = propertyType.IsValueType ? Activator.CreateInstance(propertyType) : null;
                        }
                        // Ustawienie wartości właściwości
                        SetPropertyValueByColumnName(obj, property.Name, value);
                     }
                  }
                  result.Add(obj);
               }
            }
         }
         catch (Exception ex) 
         {
            base.RaiseErrorEvent(ex);
         }
         return result; 
      }

      protected int Execute(SqlDatabaseConnection connection, string sqlSelect, NetParameters sqlParameters = null)
      {
         int result = 0;
         bool state = ((connection != null) && (connection.State == ConnectionState.Open));
         if (!state) throw new InvalidOperationException("Brak połączenia do bazy danych SQL.");

         if (!sqlSelect.IsNullOrEmpty())
         try
         {
            string sql = PrepareSQL(sqlSelect, sqlParameters);
            using (SqlDatabaseCommand command = new SqlDatabaseCommand())
            {
               command.Connection = connection;
               command.CommandText = sql;
               if (sqlParameters != null)
               {
                  foreach (var par in sqlParameters.Items)
                     command.Parameters.AddWithValue(par.ParName, par.ParValue);
               }
               result = command.ExecuteReader().RecordsAffected;
            }
         }
         catch (Exception ex)
         {
            base.RaiseErrorEvent(ex);
         }
         return result;
      }

      public int Execute(string sqlSelect, NetParameters sqlParameters = null)
      {
         return Execute(_Connection, sqlSelect, sqlParameters);
      }

      protected bool ExecuteAll(SqlDatabaseConnection connection, string SqlExecute, NetParameters sqlParameters = null)
      {
         bool result = true;
         List<string> que = new List<string>();
         que.AddRange(SqlExecute.Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries));

         try
         {
            var transaction = BeginTransaction(connection);
            if (transaction != null)
            {
               try
               {
                  foreach (string sql in que)
                  {
                     string commandSql = sql.Trim();
                     if (commandSql.IsNullOrEmpty()) continue;
                     result = Execute(connection, commandSql, sqlParameters) != -1;
                     if (!result) break;
                  }
               }
               finally
               {
                  if (result)
                     transaction.Commit();
                  else
                     transaction.Rollback();
               }
            }
         }
         catch (Exception ex)
         {
            base.RaiseErrorEvent(ex);
         }
         return result;
      }

      public bool ExecuteAll(string SqlExecute, NetParameters sqlParameters = null)
      {
         return ExecuteAll(_Connection, SqlExecute, sqlParameters);
      }

      protected object LastInseredRowId(SqlDatabaseConnection connection)
      {
         object result = null;
         bool state = ((connection != null) && (connection.State == ConnectionState.Open));
         if (!state) throw new InvalidOperationException("Brak połączenia do bazy danych SQL.");

         using (SqlDatabaseCommand command = new SqlDatabaseCommand())
         {
            command.Connection = connection;
            command.CommandText = "SELECT last_insert_rowid()";
            result = command.ExecuteScalar();
         }
         return result;
      }

      public object LastInseredRowId()
      {
         return LastInseredRowId(_Connection);
      }

      protected List<string> SQLUpdateFiles(string sqlUpdateDir, string dbVersion)
      {
         List<string> files = new List<string>();
         foreach (var file in Directory.GetFiles(sqlUpdateDir, "*.sql"))
         {
            string dbPatch = Path.GetFileNameWithoutExtension(file);
            if (string.Compare(dbPatch, dbVersion, StringComparison.OrdinalIgnoreCase)  > 0) files.Add(dbPatch);
         }
         files.Sort();
         return files;
      }


      public bool ValidatePassword(string Password, int MinLength = 5, int MinLetter = 2, int MinDigit = 1, int MinSpecjal = 0)
      {
         bool result = Password.Length >= MinLength;
         int Letter = 0;
         int Digit = 0;
         int Specjal = 0;
         foreach (char c in Password)
         {
            if (Char.IsDigit(c)) Digit++;
            else
            if (Char.IsLetter(c)) Letter++;
            else
               Specjal++;
         }
         if (Letter < MinLetter) result = false;
         if (Digit < MinDigit) result = false;
         if (Specjal < MinSpecjal) result = false;
         return result;
      }

      public bool ChangePassword(String Password)
      {
         bool result = false;
         try
         {
            bool state = ((_Connection != null) && (_Connection.State == ConnectionState.Open));
            if (!state) throw new InvalidOperationException("Brak połączenia do bazy danych SQL.");
            string fileKey = Path.ChangeExtension(Path.Combine(DirDb, DatabaseName), ".key");
            if (!File.Exists(fileKey)) throw new InvalidOperationException($"Brak pliku klucza: {fileKey}");
            string mixedKey = KeyBuilder(Password);

            using (SqlDatabaseCommand command = new SqlDatabaseCommand())
            {
               // zmiana kodowania w bazie
               command.Connection = _Connection;
               command.CommandText = $"SYSCMD ReKey = '{mixedKey}'";
               command.ExecuteScalar();

               // odczytanie klucza
               string bodyKey = File.ReadAllText(fileKey);
               string bodyKeyDecrypt = XTEA.Decrypt(bodyKey, Const.KEY_SecurePhraseFile);
               SignatureFile signature = JsonConvert.DeserializeObject<SignatureFile>(bodyKeyDecrypt) as SignatureFile;
               // zmiana hasła
               signature.HashPassword = HashPassword(Password);
               // zapis klucza
               string jsonString = JsonConvert.SerializeObject(signature);
               string bodyKeyCrypt = XTEA.Encrypt(jsonString, Const.KEY_SecurePhraseFile);
               File.WriteAllText(fileKey, bodyKeyCrypt);
               result = true;
            }
         }
         catch (Exception ex)
         {
            base.RaiseErrorEvent(ex);
         }
         return result;
      }

      public string HashPassword(string Password)
      {
         return Crc32Utils.Crc32String(Password);
      }

      protected string KeyBuilder(string Password)
      {
         string result = String.Empty;

         int size = Math.Max(_PhraseKey.Length, Password.CastAsString(string.Empty).Length);
         for (int iii=0; iii < size; iii++)
         {
            if (iii < Password.Length) result += Password[iii];
            if (iii < _PhraseKey.Length) result += _PhraseKey[iii];
         }
         return result;
      }
   }

   public class MaintainsNet : DatabaseNet
   {
      private const string DataBaseVersion = "Program::WersjaBazy";

      public MaintainsNet(string dirDb, string phraseKey) : base(dirDb, phraseKey) 
      {
      }

      public bool CreateDatabase(string NameDb, string Password, string StructSQL)
      {
         bool result = false;
         try
         {
            string fileDb = Path.ChangeExtension(Path.Combine(DirDb, NameDb), ".db");
            string fileKey = Path.ChangeExtension(Path.Combine(DirDb, NameDb), ".key");
            string mixedKey = KeyBuilder(Password);

            var connection = new SqlDatabaseConnection();
            connection.ConnectionString = $"SchemaName=db;uri=file://{fileDb}";
            connection.Open();
            if (connection.State == ConnectionState.Open)
            {
               using (SqlDatabaseCommand command = new SqlDatabaseCommand())
               {
                  command.Connection = connection;
                  command.CommandText = $"SYSCMD Key = '{mixedKey}'";
                  command.ExecuteScalar();
                  // utworz strukturę
                  if (ExecuteAll(connection, StructSQL))
                  {
                     SignatureFile signature = new SignatureFile();
                     signature.HashPassword = HashPassword(Password);
                     signature.ExpireDate = DateTime.Now.AddMonths(3);
                     signature.KeyXsec = GenerateKey.RandomString(128);

                     string jsonString = JsonConvert.SerializeObject(signature);
                     string bodyKeyCrypt = XTEA.Encrypt(jsonString, Const.KEY_SecurePhraseFile);
                     File.WriteAllText(fileKey, bodyKeyCrypt);
                     result = true;
                  }
               }
            }
            // zamknij jeżeli było otwarte
            connection?.Close();
         }
         catch (Exception ex)
         {
            base.RaiseErrorEvent(ex);
         }
         return result;
      }

      public bool UpdateDatabase(string NameDb, string Password, string sqlUpdateDir)
      {
         bool result = false;

         string fileDb = Path.ChangeExtension(Path.Combine(DirDb, NameDb), ".db");
         string fileKey = Path.ChangeExtension(Path.Combine(DirDb, NameDb), ".key");
         if (!File.Exists(fileDb)) throw new InvalidOperationException($"Brak pliku bazy danych: {fileDb}");
         if (!File.Exists(fileKey)) throw new InvalidOperationException($"Brak pliku klucza: {fileKey}");

            string bodyKey = File.ReadAllText(fileKey);
            string bodyKeyDecrypt = XTEA.Decrypt(bodyKey, Const.KEY_SecurePhraseFile);
            SignatureFile signature = JsonConvert.DeserializeObject<SignatureFile>(bodyKeyDecrypt) as SignatureFile;

         if (signature.HashPassword == HashPassword(Password))
         {
            string mixedKey = KeyBuilder(Password);
            try
            {
               var connection = new SqlDatabaseConnection();
               connection.ConnectionString = $"SchemaName=db;uri=file://{fileDb}";
               connection.Open();
               if (connection.State == ConnectionState.Open)
               {
                  using (SqlDatabaseCommand command = new SqlDatabaseCommand())
                  {
                     command.Connection = connection;
                     command.CommandText = $"SYSCMD Key = '{mixedKey}'";
                     command.ExecuteScalar();
                     // jestem połączony, pobierz wersję bazy
                     string actVersion = base.SelectScalar<string>(connection, $"SELECT [WartoscS] FROM [Parametry] WHERE [Klucz]='{DataBaseVersion}'");
                     foreach (string updVersion in base.SQLUpdateFiles(sqlUpdateDir, actVersion))
                     {
                        string sqlUpdateFile = Path.Combine(sqlUpdateDir, $"{updVersion}.sql");
                        if (File.Exists(sqlUpdateFile))
                        {
                           string sqlUpdateText = File.ReadAllText(sqlUpdateFile);
                           if (!base.ExecuteAll(connection, sqlUpdateText))
                              break;
                        }
                     }
                     result = true;
                  }
               }
               connection?.Close();
            }
            catch (Exception ex)
            {
               base.RaiseErrorEvent(ex);
            }
         }
         return result;
      }

   } 

   public class SetTable : MyClass
   {
      public Type DataType;
      public string TableName;
      public List<ColumnInfo> Columns = new List<ColumnInfo>();
      public string SQLSelect { get; set; }
      public string SQLInsert { get; set; }
      public string SQLUpdate { get; set; }
      public string SQLDelete { get; set; }

      public int CountChanged = 0;

      public SetTable (string tablename, Type type)
      {
         DataType = type;
         TableName = tablename;
         ReadAttribute(type);
         SQLSelect = SQLSelectTemplate();
         SQLInsert = SQLInsertTemplate();
         SQLUpdate = SQLUpdateTemplate();
         SQLDelete = SQLDeleteTemplate();
      }

      private void ReadAttribute(Type type)
      {
         // Odczytaj atrybuty za pomocą refleksji
         foreach (var property in type.GetProperties())
         {
            ColumnAttribute _ColumnAttribute = null;
            PrimaryKeyAttribute _PrimaryKeyAttribute = null;
            ForeignKeyAttribute _ForeignKeyAttribute = null;

            // przez wszystkie atrybuty kolumny
            foreach (var attribute in property.GetCustomAttributes())
            {
               if (attribute is ColumnAttribute columnAttribute)
                  _ColumnAttribute = columnAttribute;
               if (attribute is PrimaryKeyAttribute primaryKeyAttribute)
                  _PrimaryKeyAttribute = primaryKeyAttribute;
               if (attribute is ForeignKeyAttribute foreignKeyAttribute)
                  _ForeignKeyAttribute = foreignKeyAttribute;
            }

            // zapisz
            if (_ColumnAttribute != null)
            {
               ColumnInfo c = new ColumnInfo();
               c.ColumnName = _ColumnAttribute.Name;
               c.TableColumn = !_ColumnAttribute.VirtualColumn;
               c.PrimaryKey = _PrimaryKeyAttribute != null;
               c.AutoIncremental = _PrimaryKeyAttribute != null && _PrimaryKeyAttribute.AutoIncrement;
               c.ForeignKey = _ForeignKeyAttribute != null;
               c.ReferencedTable = _ForeignKeyAttribute?.ReferencedTable ?? string.Empty;
               c.ReferencedColumn = _ForeignKeyAttribute?.ReferencedColumn ?? string.Empty;
               Columns.Add(c);
            }
         }
      }

      private string SQLSelectTemplate()
      {
         var columns = GetColumnNames(true);
         return $"SELECT {string.Join(", ", columns)} FROM {TableName} WHERE {SQLWherePrimaryKey()}";
      }

      private string SQLSelectWhereTemplate(string[] columnsWhere)
      {
         var columns = GetColumnNames(true);
         string Where = string.Empty;
         foreach (var column in columnsWhere)
            Where += $" AND {column}=@{column}";
         return $"SELECT {string.Join(", ", columns)} FROM {TableName} WHERE (1=1) {Where}";
      }

      private string SQLInsertTemplate()
      {
         var columns = GetColumnNames(false);
         var parameters = GetParameterNames(false);
         return $"INSERT INTO {TableName} ({string.Join(", ", columns)}) VALUES ({string.Join(", ", parameters)})";
      }

      private string SQLUpdateTemplate()
      {
         var setColumns = Columns.Where(w => w.TableColumn && !w.PrimaryKey).Select(c => $"[{c.ColumnName}] = @{c.ColumnName}");
         return $"UPDATE {TableName} SET {string.Join(", ", setColumns)} WHERE {SQLWherePrimaryKey()}";
      }

      private string SQLDeleteTemplate()
      {
         string SqlListColumn = string.Empty;
         foreach (ColumnInfo item in Columns.Where(w => w.PrimaryKey == false))
         {
            if (!String.IsNullOrEmpty(SqlListColumn)) SqlListColumn += ", ";
            SqlListColumn += $"[{item.ColumnName}] = @{item.ColumnName}";
         }
         return $"DELETE FROM {TableName} WHERE {SQLWherePrimaryKey()}";
      }

      private string SQLWherePrimaryKey()
      {
         string sqlCommand = string.Empty;
         foreach (ColumnInfo item in Columns.Where(w => w.PrimaryKey == true))
         {
            if (!String.IsNullOrEmpty(sqlCommand)) sqlCommand += " AND ";
            sqlCommand += $"[{item.ColumnName}] = @{item.ColumnName}";
         }
         return sqlCommand;
      }

      private IEnumerable<string> GetColumnNames(bool includePrimaryKey)
      {
         return Columns.Where(c => c.TableColumn && (!c.AutoIncremental || includePrimaryKey)).Select(c => "["+c.ColumnName+"]");
      }

      private IEnumerable<string> GetParameterNames(bool includePrimaryKey)
      {
         return Columns.Where(c => c.TableColumn &&  (!c.AutoIncremental || includePrimaryKey)).Select(c => "@" + c.ColumnName);
      }

      public bool NewRecord<T>(T result) where T : class
      {
         bool success = false;
         try
         {
            // parametry primary key (bez autoincremental)
            foreach (ColumnInfo column in Columns)
            {
               var property = typeof(T).GetProperty(column.ColumnName);
               if (property != null)
               {
                  // Typ właściwości (obsługa nullable i typów prostych)
                  var propertyType = Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType;

                  // Ustawienie domyślnej wartości dla typów prostych
                  object value = propertyType.IsValueType ? Activator.CreateInstance(propertyType) : null;

                  // Ustawienie wartości właściwości
                  DatabaseNet.SetPropertyValueByColumnName(result, property.Name, value);
               }
            }
            success = true;
         }
         catch (Exception ex) 
         {
            base.RaiseErrorEvent(ex);
         }
         return success;
      }

      public bool Single<T>(SqlDatabaseConnection connection, T result) where T : class
      {
         bool success = false; 
         try
         {
            if ((connection != null) && (connection.State == ConnectionState.Open))
            {
               using (SqlDatabaseCommand command = new SqlDatabaseCommand())
               {
                  command.Connection = connection;
                  command.CommandText = SQLSelect;

                  // parametry primary key
                  foreach (ColumnInfo column in Columns.Where(c => c.TableColumn && c.PrimaryKey))
                  {
                     var property = typeof(T).GetProperty(column.ColumnName);
                     if (property != null)
                     {
                        var value = property.GetValue(result);
                        if (value == null) 
                           throw new InvalidOperationException($"Wartość dla kolumny {column.ColumnName} jest null.");
                        command.Parameters.AddWithValue("@" + column.ColumnName, value);
                     }
                  }
                  
                  using (var reader = command.ExecuteReader())
                  {
                     if (reader.Read())  // Tylko jeden rekord
                     {
                        success = true;
                        foreach (ColumnInfo column in Columns)
                        {
                           var property = typeof(T).GetProperty(column.ColumnName);

                           if (property != null && column.TableColumn)
                           {
                              var value = reader[column.ColumnName];
                              if (value != DBNull.Value)
                                 property.SetValue(result, value);
                           }
                        }
                     }
                  }

               }
            }
         }
         catch (Exception ex)
         {
            RaiseErrorEvent(ex);
         }
         return success;
      }

      public bool SingleWhere<T>(SqlDatabaseConnection connection, T result, string[] columnsWhere) where T : class
      {
         bool success = false;
         try
         {
            if ((connection != null) && (connection.State == ConnectionState.Open))
            {
               using (SqlDatabaseCommand command = new SqlDatabaseCommand())
               {
                  command.Connection = connection;
                  command.CommandText = SQLSelectWhereTemplate(columnsWhere);

                  // parametry primary key
                  foreach (ColumnInfo column in Columns.Where(c => c.TableColumn))
                  {
                     if (command.CommandText.IndexOf($"@{column.ColumnName}", StringComparison.OrdinalIgnoreCase) >= 0)
                     {
                        var property = typeof(T).GetProperty(column.ColumnName);
                        if (property != null)
                        {
                           var value = property.GetValue(result);
                           if (value == null)
                              throw new InvalidOperationException($"Wartość dla kolumny {column.ColumnName} jest null.");
                           command.Parameters.AddWithValue("@" + column.ColumnName, value);
                        }
                     }
                  }

                  using (var reader = command.ExecuteReader())
                  {
                     if (reader.Read())  // Tylko jeden rekord
                     {
                        success = true;
                        foreach (ColumnInfo column in Columns)
                        {
                           var property = typeof(T).GetProperty(column.ColumnName);
                           if (property != null && column.TableColumn)
                           {
                              var value = reader[column.ColumnName];
                              if (value != DBNull.Value)
                                 property.SetValue(result, value);
                           }
                        }
                     }
                  }

               }
            }
         }
         catch (Exception ex)
         {
            RaiseErrorEvent(ex);
         }

         return success;
      }

      public bool Update<T>(SqlDatabaseConnection connection, T result) where T : class
      {
         bool success = false;
         try
         {
            if ((connection != null) && (connection.State == ConnectionState.Open))
            {
               using (SqlDatabaseCommand command = new SqlDatabaseCommand())
               {
                  command.Connection = connection;
                  command.CommandText = SQLUpdate;

                  // parametry primary key
                  foreach (ColumnInfo column in Columns.Where(c => c.TableColumn && c.PrimaryKey))
                  {
                     var property = typeof(T).GetProperty(column.ColumnName);
                     if (property != null)
                     {
                        var value = property.GetValue(result);
                        if (value == null)
                           throw new InvalidOperationException($"Wartość dla kolumny {column.ColumnName} jest null.");
                        command.Parameters.AddWithValue("@" + column.ColumnName, value);
                     }
                  }
                  // parametry - wartość pól
                  foreach (ColumnInfo column in Columns.Where(c => c.TableColumn && !c.PrimaryKey))
                  {
                     var property = typeof(T).GetProperty(column.ColumnName);
                     if (property != null)
                     {
                        var value = property.GetValue(result);
                        command.Parameters.AddWithValue("@" + column.ColumnName, value);
                     }
                  }
                  success = command.ExecuteNonQuery() == 1;
                  if (success) CountChanged++;
               }
            }
         }
         catch (Exception ex)
         {
            RaiseErrorEvent(ex);
         }
         return success;
      }

      public bool Insert<T>(SqlDatabaseConnection connection, T result) where T : class
      {
         bool success = false;
         try
         {
            if ((connection != null) && (connection.State == ConnectionState.Open))
            {
               using (SqlDatabaseCommand command = new SqlDatabaseCommand())
               {
                  command.Connection = connection;
                  command.CommandText = SQLInsert;

                  // parametry primary key (bez autoincremental)
                  foreach (ColumnInfo column in Columns.Where(c => c.TableColumn && (c.PrimaryKey && !c.AutoIncremental)))
                  {
                     var property = typeof(T).GetProperty(column.ColumnName);
                     if (property != null)
                     {
                        var value = property.GetValue(result);
                        if (value == null)
                           throw new InvalidOperationException($"Wartość dla kolumny {column.ColumnName} jest null.");
                        command.Parameters.AddWithValue("@" + column.ColumnName, value);
                     }
                  }
                  // parametry primary key (dla autoincremental)
                  foreach (ColumnInfo column in Columns.Where(c => c.TableColumn && (c.PrimaryKey && c.AutoIncremental)))
                  {
                     var property = typeof(T).GetProperty(column.ColumnName);
                     if (property != null)
                     {
                        var propertyType = property.PropertyType;

                        if ( propertyType == typeof(int) || propertyType == typeof(long) )
                          property.SetValue(result, 0);
                     }
                  }

                  // parametry - wartość pól
                  foreach (ColumnInfo column in Columns.Where(c => c.TableColumn && !c.PrimaryKey))
                  {
                     var property = typeof(T).GetProperty(column.ColumnName);
                     if (property != null)
                     {
                        var value = property.GetValue(result);
                        command.Parameters.AddWithValue("@" + column.ColumnName, value);
                     }
                  }
                  success = command.ExecuteNonQuery() == 1;
                  if (success)
                  {
                     var pk = Columns.Where(c => c.TableColumn && (c.PrimaryKey && c.AutoIncremental));
                     if (pk.Count() == 1)
                     {
                        command.CommandText = "SELECT last_insert_rowid()";
                        object rowid = command.ExecuteScalar();
                        var property = typeof(T).GetProperty(pk.First().ColumnName);
                        if (property != null)
                        {
                           var propertyType = property.PropertyType;

                           if (propertyType == typeof(int) || propertyType == typeof(long))
                              property.SetValue(result, Convert.ChangeType(rowid, propertyType));
                        }
                     }
                     CountChanged++;
                  }

               }
            }
         }
         catch (Exception ex)
         {
            RaiseErrorEvent(ex);
         }
         return success;
      }

      public bool Delete<T>(SqlDatabaseConnection connection, T result) where T : class
      {
         bool success = false;
         try
         {
            if ((connection != null) && (connection.State == ConnectionState.Open))
            {
               using (SqlDatabaseCommand command = new SqlDatabaseCommand())
               {
                  command.Connection = connection;
                  command.CommandText = SQLDelete;

                  // parametry primary key
                  foreach (ColumnInfo column in Columns.Where(c => c.TableColumn && c.PrimaryKey))
                  {
                     var property = typeof(T).GetProperty(column.ColumnName);
                     if (property != null)
                     {
                        var value = property.GetValue(result);
                        if (value == null)
                           throw new InvalidOperationException($"Wartość dla kolumny {column.ColumnName} jest null.");
                        command.Parameters.AddWithValue("@" + column.ColumnName, value);
                     }
                  }
                  success = command.ExecuteNonQuery() == 1;
               }
            }
         }
         catch (Exception ex)
         {
            RaiseErrorEvent(ex);
         }
         return success;
      }


      public class ColumnInfo
      {
         public string ColumnName { get; set; }
         public bool TableColumn { get; set; }
         public bool PrimaryKey { get; set; }
         public bool AutoIncremental { get; set; }
         public bool ForeignKey { get; set; }
         public string ReferencedTable { get; set; }
         public string ReferencedColumn { get; set; }
      }
   }

   public class NetParameters
   {
      List<ParameterInfo> items = new List<ParameterInfo>();

      public List<ParameterInfo> Items { get { return items;  }  }

      public void Add(string parName, object parValue)
      {
         if (parName.IsNullOrEmpty()) return;

         var item = Items.Where(x => x.ParName == parName).SingleOrDefault();
         if (item == null)
            items.Add(new ParameterInfo() { ParName = parName, ParValue = parValue });
         else
            item.ParValue = parValue;
      }

      public int Count
      {
         get { return Items.Count(); }
      }

      public class ParameterInfo
      {
         public string ParName { get; set; }
         public object ParValue { get; set; }
      }
   }


   [AttributeUsage(AttributeTargets.Property)]
   public class ColumnAttribute : Attribute
   {
      public string Name { get; }
      public bool VirtualColumn { get; set; } = false;

      public ColumnAttribute(string name)
      {
         Name = name;
      }
   }

   [AttributeUsage(AttributeTargets.Property)]
   public class PrimaryKeyAttribute : Attribute
   {
      public bool AutoIncrement { get; set; } = true;

      public PrimaryKeyAttribute() { }
      public PrimaryKeyAttribute(bool autoIncrement)
      {
         AutoIncrement = autoIncrement;
      }
   }

   [AttributeUsage(AttributeTargets.Property)]
   public class ForeignKeyAttribute : Attribute
   {
      public string ReferencedTable { get; }
      public string ReferencedColumn { get; }

      public ForeignKeyAttribute(string referencedTable, string referencedColumn)
      {
         ReferencedTable = referencedTable;
         ReferencedColumn = referencedColumn;
      }
   }

   public static class DbUtils
   {
      public static Bitmap ByteToImage(byte[] arrByte)
      {
         try
         {
            using (MemoryStream mStream = new MemoryStream())
            {
               byte[] pData = arrByte;
               mStream.Write(pData, 0, Convert.ToInt32(pData.Length));
               Bitmap bm = new Bitmap(mStream, false);
               return bm;
            }
         }
         catch
         {
            return null;
         }
      }

   }

   public class Dict_IdName
   {
      public int Id { get; set; }
      public string Name { get; set; }
   }
}

