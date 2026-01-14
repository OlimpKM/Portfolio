-- -------------------------------------------------------------------------------------
-- Financial Data Aggregator
-- Description: Aggregates calculated amounts, interest, and costs across multiple
-- currencies for attachment records.
-- -------------------------------------------------------------------------------------

-- Cleanup temporary processing table
DELETE FROM TMP_ATTACHMENT_CURRENCY;

-- Insert aggregated financial metrics
INSERT INTO TMP_ATTACHMENT_CURRENCY (
    `ATTACHMENT_ID`, `CALC_AMOUNT`, `CALC_INTEREST`, `CURRENCY_ID`,
    `PAID_AMOUNT`, `PAID_INTEREST`, `PAID_COST`, `PAID_OTHER`, `PAID_TOTAL`,
    `BOOK_AMOUNT`, `BOOK_INTEREST`, `BOOK_COST`, `BOOK_OTHER`, `BOOK_TOTAL`,
    `CURRENCY_CODE`, `VALUATION_CREDITOR`, `VALUATION_COMPANY`
)
SELECT
    T.`ATTACHMENT_ID`,
    IFNULL(W.`CALC_AMOUNT`, 0) AS 'CALC_AMOUNT',
    IFNULL(W.`CALC_INTEREST`, 0) AS 'CALC_INTEREST',
    T.`CURRENCY_ID`,
    -- Payments Breakdown
    IFNULL(W.`PAID_AMOUNT`, 0),
    IFNULL(W.`PAID_INTEREST`, 0),
    IFNULL(W.`PAID_COST`, 0),
    IFNULL(W.`PAID_OTHER`, 0),
    IFNULL(W.`PAID_TOTAL`, 0),
    -- Booking/Accounting Breakdown
    IFNULL(W.`BOOK_AMOUNT`, 0),
    IFNULL(W.`BOOK_INTEREST`, 0),
    IFNULL(W.`BOOK_COST`, 0),
    IFNULL(W.`BOOK_OTHER`, 0),
    IFNULL(W.`BOOK_TOTAL`, 0),
    T.`CURRENCY_SYMBOL`,
    T.`VAL_CREDITOR_RATE`,
    T.`VAL_COMPANY_RATE`
FROM (
    SELECT
        ATT.`ID_ATTACHMENT` AS `ATTACHMENT_ID`,
        ATT.`ID_CURRENCY` AS `CURRENCY_ID`,
        CUR.`SYMBOL` AS `CURRENCY_SYMBOL`,
        ATT.`RATE_CREDITOR` AS `VAL_CREDITOR_RATE`,
        ATT.`RATE_COMPANY` AS `VAL_COMPANY_RATE`
    FROM `ATTACHMENTS` AS ATT
    LEFT JOIN `CURRENCIES` AS CUR ON (ATT.`ID_CURRENCY` = CUR.`ID_CURRENCY`)
    WHERE ATT.`IS_ACTIVE` = 1
) AS T
LEFT JOIN (
    -- Subquery for calculating actual sums from history/transactions
    SELECT
        HIST.`ID_ATTACHMENT` AS `ATTACHMENT_ID`,
        SUM(HIST.`AMOUNT`) AS `PAID_AMOUNT`,
        SUM(HIST.`INTEREST`) AS `PAID_INTEREST`,
        SUM(HIST.`COSTS`) AS `PAID_COST`,
        -- ... additional aggregation logic ...
        MAX(HIST.`LAST_UPDATE`) AS `LATEST_TRANS`
    FROM `TRANSACTION_HISTORY` AS HIST
    GROUP BY HIST.`ID_ATTACHMENT`
) AS W ON (T.`ATTACHMENT_ID` = W.`ATTACHMENT_ID`);

-- -------------------------------------------------------------------------------------
-- Final view logic for reporting
-- -------------------------------------------------------------------------------------
SELECT
    RES.*,
    CTR.`CONTRACT_NUMBER`,
    CLI.`CLIENT_NAME` AS `DEBTOR`,
    CRD.`CLIENT_NAME` AS `CREDITOR`
FROM TMP_ATTACHMENT_CURRENCY AS RES
INNER JOIN `ATTACHMENTS` AS ATT ON (RES.`ATTACHMENT_ID` = ATT.`ID_ATTACHMENT`)
LEFT JOIN `CONTRACTS` AS CTR ON (ATT.`ID_CONTRACT` = CTR.`ID_CONTRACT`)
LEFT JOIN `CLIENTS` AS CLI ON (ATT.`ID_DEBTOR` = CLI.`ID_CLIENT`)
LEFT JOIN `CLIENTS` AS CRD ON (CTR.`ID_CREDITOR` = CRD.`ID_CLIENT`)
ORDER BY RES.`ATTACHMENT_ID`;