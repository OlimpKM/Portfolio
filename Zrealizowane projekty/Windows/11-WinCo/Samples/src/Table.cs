using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using Dapper;
using System.Linq;

namespace WinCo.Core.Repositories
{
    /// <summary>
    /// Repozytorium odpowiedzialne za obs³ugê wiadomoœci przychodz¹cych.
    /// Przyk³ad implementacji bezpiecznego dostêpu do danych.
    /// </summary>
    public class MessageRepository
    {
        private readonly string _connectionString;

        public MessageRepository(string connectionString)
        {
            _connectionString = connectionString;
        }

        /// <summary>
        /// Pobiera wiadomoœci, które nie zosta³y jeszcze rozpoznane przez automat (Parser).
        /// </summary>
        public IEnumerable<IncomingMessage> GetUnrecognizedMessages()
        {
            const string sql = @"
                SELECT Id, ReceivedAt, SenderAddress, Subject, WorkflowStatus
                FROM dms.IncomingMessages
                WHERE WorkflowStatus = @Status
                AND IsArchived = 0
                ORDER BY ReceivedAt DESC";

            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                // U¿ycie parametrów zapobiega atakom SQL Injection
                return db.Query<IncomingMessage>(sql, new { Status = 0 }).ToList();
            }
        }

        /// <summary>
        /// Przypisuje wiadomoœæ do konkretnego klienta (akcja operatora).
        /// </summary>
        public bool AssignToClient(long messageId, string clientId, string operatorCode)
        {
            const string sql = @"
                UPDATE dms.IncomingMessages
                SET ClientInternalId = @ClientId,
                    WorkflowStatus = 1,
                    OperatorCode = @OperatorCode,
                    ProcessedAt = GETDATE()
                WHERE Id = @MessageId";

            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                int rowsAffected = db.Execute(sql, new
                {
                    MessageId = messageId,
                    ClientId = clientId,
                    OperatorCode = operatorCode
                });

                return rowsAffected > 0;
            }
        }
    }

    // Model danych (POCO)
    public class IncomingMessage
    {
        public long Id { get; set; }
        public DateTime ReceivedAt { get; set; }
        public string SenderAddress { get; set; }
        public string Subject { get; set; }
        public int WorkflowStatus { get; set; }
    }
}
