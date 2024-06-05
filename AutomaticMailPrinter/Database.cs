using System;
using System.Data;
using System.Data.SQLite;

namespace AutomaticMailPrinter
{ 
    public class Database
    {
        private readonly string connectionString;

        public Database()
        {
            string databaseFile = "orders.sqlite";
            connectionString = $"Data Source={databaseFile}";

            if (!System.IO.File.Exists(databaseFile))
            {
                SQLiteConnection.CreateFile(databaseFile);
            }

            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                string createTableQuery = @"
                    CREATE TABLE IF NOT EXISTS orders (
                        id INTEGER PRIMARY KEY,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        printed_at TIMESTAMP,
                        subject TEXT,
                        html TEXT
                    )";
                using (var command = new SQLiteCommand(createTableQuery, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }

        public DataTable GetOrders(int limit, int pageSize)
        {
            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                string query = @"
                    SELECT * FROM orders
                    ORDER BY created_at DESC
                    LIMIT @Limit OFFSET @Offset";
                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Limit", limit);
                    command.Parameters.AddWithValue("@Offset", (pageSize - 1) * limit);

                    using (var adapter = new SQLiteDataAdapter(command))
                    {
                        var dataTable = new DataTable();
                        adapter.Fill(dataTable);
                        return dataTable;
                    }
                }
            }
        }

        public Order GetOrder(int id)
        {
            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                string query = @"
                SELECT * FROM orders
                WHERE id = @Id";
                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Id", id);
                    using (var reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            return new Order()
                            {
                                id = int.Parse(reader["id"].ToString()),
                                subject = reader["subject"].ToString(),
                                html = reader["html"].ToString(),
                                createdAt = (DateTime)reader["created_at"],
                                printedAt = reader["printed_at"] != null ? (DateTime)reader["printed_at"] : DateTime.MinValue,
                            };
                        }
                        else
                        {
                            return null;
                        }
                    }
                }
            }
        }

        public void AddOrder(int id, string html, string subject)
        {
            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                string query = @"
                    INSERT INTO orders (id, created_at, printed_at, html, subject)
                    VALUES (@Id, CURRENT_TIMESTAMP, NULL, @Html, @Subject)";
                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Id", id);
                    command.Parameters.AddWithValue("@Html", html);
                    command.Parameters.AddWithValue("@Subject", subject);
                    command.ExecuteNonQuery();
                }
            }
        }

        public void OrderPrinted(int id)
        {
            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                string query = @"
                    UPDATE orders
                    SET printed_at = CURRENT_TIMESTAMP
                    WHERE id = @Id";
                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Id", id);
                    command.ExecuteNonQuery();
                }
            }
        }
    }
}