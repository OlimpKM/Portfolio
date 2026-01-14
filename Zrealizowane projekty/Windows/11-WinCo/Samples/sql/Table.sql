/* Projekt: Debt Management System (DMS)
  Opis: Implementacja struktury bazy danych dla moduu komunikacji i kontroli transakcji.
*/

-- 1. Inicjalizacja dedykowanego schematu
IF NOT EXISTS (SELECT schema_name FROM information_schema.schemata WHERE schema_name = 'dms')
BEGIN
  EXEC sp_executesql N'CREATE SCHEMA dms'
END
GO

-- 2. Tabela wiadomoci przychodzcych
-- Przechowuje zanonimizowane metadane korespondencji
IF OBJECT_ID('dms.IncomingMessages') IS NULL
BEGIN
  CREATE TABLE [dms].[IncomingMessages]
  (
    [Id] [bigint] IDENTITY(1,1) NOT NULL,
    [ReceivedAt] [datetime] NOT NULL DEFAULT GETDATE(),
    [SenderAddress] [varchar](255) NOT NULL,
    [ReceiverAddress] [varchar](255) NOT NULL,
    [Subject] [nvarchar](MAX) NULL,
    [MessageBody] [nvarchar](MAX) NULL,
    [ExternalMessageId] [varchar](50) NULL, -- ID z serwera pocztowego
    [ClientInternalId] [varchar](70) NULL,
    [WorkflowStatus] [smallint] NOT NULL DEFAULT 0,
    [IsArchived] [bit] NOT NULL DEFAULT 0,
    [Checksum] [varchar](10) NULL, -- Weryfikacja integralnoci (CRC32)

    CONSTRAINT [PK_IncomingMessages_Id] PRIMARY KEY CLUSTERED ([Id] ASC)
  )

  -- Optymalizacja: Indeksy filtrowane dla unikalnych identyfikatorów zewntrznych
  CREATE UNIQUE NONCLUSTERED INDEX [IUX_IncomingMessages_ExternalId]
    ON dms.IncomingMessages([ExternalMessageId])
    WHERE [ExternalMessageId] IS NOT NULL;

  CREATE UNIQUE NONCLUSTERED INDEX [IUX_IncomingMessages_Checksum]
    ON dms.IncomingMessages([Checksum])
    WHERE [Checksum] IS NOT NULL;
END
GO

-- 3. Trigger kontrolujcy spójno transakcji
-- Zapobiega dublowaniu aktywnych rekordów o tej samej nazwie/kluczu
IF OBJECT_ID('dms.trg_CheckTransactionConsistency') IS NOT NULL
    DROP TRIGGER [dms].[trg_CheckTransactionConsistency]
GO

CREATE TRIGGER [dms].[trg_CheckTransactionConsistency] ON [dms].[IncomingMessages]
AFTER INSERT, UPDATE
AS
BEGIN
  SET NOCOUNT ON;
  -- Uycie izolacji READ UNCOMMITTED tylko dla sprawdzenia istnienia duplikatów
  SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;

  DECLARE @TransactionName varchar(50);
  DECLARE @ActiveCount int;

  -- Pobranie nazwy z nowo wstawionego rekordu
  SELECT @TransactionName = i.ExternalMessageId FROM inserted i;

  -- Sprawdzenie czy istnieje ju aktywny rekord o tym samym identyfikatorze
  SELECT @ActiveCount = COUNT(Id)
  FROM dms.IncomingMessages
  WHERE ExternalMessageId = @TransactionName;

  IF @ActiveCount > 1
  BEGIN
    RAISERROR ('Bd: Wykryto duplikacj aktywnej transakcji w systemie.', 16, 1);
    ROLLBACK TRANSACTION;
  END
END
GO