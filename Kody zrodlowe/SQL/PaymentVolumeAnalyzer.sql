-- -------------------------------------------------------------------------------------
-- Payment Volume Analyzer
-- Description: Calculates the total sum of payments for a filtered set of assignments
-- within a specific date range.
-- -------------------------------------------------------------------------------------

SELECT
    SUM(P.`AMOUNT`) AS 'TOTAL_PAYMENTS_VOLUME'
FROM `PAYMENTS` AS P
LEFT OUTER JOIN `ASSIGNMENTS` AS A ON (A.`ID_ASSIGNMENT` = P.`ID_ASSIGNMENT`)
INNER JOIN
(
    -- [SUBQUERY: Defines the scope of active assignments based on business criteria]
    SELECT
        A_SUB.`ID_ASSIGNMENT`
    FROM `ASSIGNMENTS` AS A_SUB
    LEFT OUTER JOIN `CONTRACTS` AS C ON (C.`ID_CONTRACT` = A_SUB.`ID_CONTRACT`)
    LEFT OUTER JOIN `CLIENTS` AS D ON (D.`ID_CLIENT` = A_SUB.`ID_DEBTOR`)
    LEFT OUTER JOIN `CLIENTS` AS CR ON (CR.`ID_CLIENT` = C.`ID_CREDITOR`)
    LEFT OUTER JOIN `COLLECTION_SERVICES` AS CS ON (A_SUB.`ID_SERVICE` = CS.`ID_SERVICE`)
    -- Placeholder for dynamic business logic filters (e.g., specific operator or status)
    WHERE {-DYNAMIC_BUSINESS_FILTERS-}
) AS _SCOPE ON (_SCOPE.`ID_ASSIGNMENT` = P.`ID_ASSIGNMENT`)
WHERE
    (A.`ASSIGNMENT_DATE` >= :DATE_FROM)
    AND
    (A.`ASSIGNMENT_DATE` <= :DATE_TO);