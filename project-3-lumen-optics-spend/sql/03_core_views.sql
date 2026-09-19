/*
================================================================================
Project 3 -- Lumen Optics Manufacturing: Photonics Spend Scorecard
Script:  03_core_views.sql
Purpose: THE single definition of what a purchase actually cost, and what price
         it should have been charged at. Every downstream object -- quality
         checks, KPI views, the priority queue, the DAX measures and the Excel
         formulas -- reads from here, so one definition can never drift into
         several.

THE LANDED COST IDENTITY (the formula everything else depends on):

    LandedCost = OrderQty x UnitPrice
               + FreightAmount
               + ExpediteFee

    CostPerAcceptedUnit = LandedCost / QtyAccepted

The second line is the one that changes decisions. Unit price answers "what did
we agree to pay"; cost per ACCEPTED unit answers "what did each usable part
cost". In optics those differ enough to reverse a supplier ranking: a component
that arrives out of specification was still paid for, still freighted, and
cannot be built into a module.

WHY THE CONTRACTED PRICE IS RESOLVED BY A FUNCTION, NOT A JOIN
    An agreement is true between two dates. Resolving it needs the agreement in
    force at the ORDER date, treating an open ValidTo as "still in force" --
    and a plain join gets that wrong in two directions at once: it drops
    open-ended agreements when the predicate is written carelessly, and it
    multiplies rows when two agreements overlap.

    Overlaps are real in this dataset (they are planted), so the resolution is
    deterministic by construction -- latest start wins, then highest key -- and
    it also returns AgreementsInForce so the quality layer can report that the
    contracted price was ambiguous rather than silently picking one.

DATA DISCLOSURE: Lumen Optics Manufacturing is fictional; all data is synthetic.
================================================================================
*/

USE LumenSpend;
GO

IF OBJECT_ID('dbo.vw_POLineCost', 'V')      IS NOT NULL DROP VIEW dbo.vw_POLineCost;
IF OBJECT_ID('dbo.fn_POLineCost', 'IF')     IS NOT NULL DROP FUNCTION dbo.fn_POLineCost;
IF OBJECT_ID('dbo.fn_ContractedPrice', 'IF') IS NOT NULL DROP FUNCTION dbo.fn_ContractedPrice;
GO

/*
--------------------------------------------------------------------------------
fn_ContractedPrice(@PartKey, @VendorKey, @OrderDateKey)

Returns the agreement in force at the order date, or nothing at all when the
purchase was off contract. AgreementsInForce is returned alongside the price
because "there were two" is a different problem from "there was none", and both
are invisible if the function returns only a number.
--------------------------------------------------------------------------------
*/
CREATE FUNCTION dbo.fn_ContractedPrice (@PartKey INT, @VendorKey INT, @OrderDateKey INT)
RETURNS TABLE
AS RETURN
(
    SELECT TOP 1
        pa.AgreementKey,
        pa.AgreementNo,
        pa.AgreedUnitPrice,
        pa.MinOrderQty,
        -- how many agreements covered this order date. More than one means the
        -- contracted price is ambiguous and the variance below is only as good
        -- as the tie-break.
        AgreementsInForce = (
            SELECT COUNT(*)
            FROM dbo.Fact_PriceAgreement p2
            WHERE p2.PartKey = @PartKey AND p2.VendorKey = @VendorKey
              AND p2.ValidFromDateKey <= @OrderDateKey
              AND (p2.ValidToDateKey IS NULL OR p2.ValidToDateKey >= @OrderDateKey))
    FROM dbo.Fact_PriceAgreement pa
    WHERE pa.PartKey = @PartKey AND pa.VendorKey = @VendorKey
      AND pa.ValidFromDateKey <= @OrderDateKey
      -- an open-ended agreement is still in force; excluding NULL here is the
      -- usual way contract coverage is made to look far worse than it is
      AND (pa.ValidToDateKey IS NULL OR pa.ValidToDateKey >= @OrderDateKey)
    -- deterministic tie-break: latest start, then highest key. Without the
    -- second term two overlapping agreements return whichever row the engine
    -- happens to reach first, and the same query gives different answers.
    ORDER BY pa.ValidFromDateKey DESC, pa.AgreementKey DESC
);
GO

/*
--------------------------------------------------------------------------------
fn_POLineCost(@AsOf) -- one row per purchase order line, carrying the landed
cost, the contracted position, and what was actually usable.

Receipts are aggregated to the line because a line may be delivered in several
drops. Lines with no receipt yet are kept, with NULL acceptance, rather than
dropped: an order placed and not yet delivered is a real commitment, and
silently excluding it understates both spend and exposure.
--------------------------------------------------------------------------------
*/
CREATE FUNCTION dbo.fn_POLineCost (@AsOf DATE)
RETURNS TABLE
AS RETURN
(
    WITH AsOfKey AS (SELECT k = YEAR(@AsOf)*10000 + MONTH(@AsOf)*100 + DAY(@AsOf)),
    Rcpt AS (
        SELECT g.POLineKey,
               QtyReceived = SUM(g.QtyReceived),
               QtyAccepted = SUM(g.QtyAccepted),
               QtyRejected = SUM(g.QtyRejected),
               Receipts    = COUNT(*),
               FirstReceiptKey = MIN(g.ReceiptDateKey),
               LastReceiptKey  = MAX(g.ReceiptDateKey)
        FROM dbo.Fact_GoodsReceipt g
        CROSS JOIN AsOfKey a
        WHERE g.ReceiptDateKey <= a.k
        GROUP BY g.POLineKey
    )
    SELECT
        l.POLineKey, l.PONumber, l.POLineNo,
        l.PartKey, l.VendorKey, l.BuyerKey,
        OrderDate    = do.[Date],
        PromisedDate = dp.[Date],
        OrderYear    = do.[Year],
        OrderYearMonth = do.YearMonth,
        l.OrderQty, l.UnitPrice, l.FreightAmount, l.ExpediteFee, l.IsExpedited,

        -- the landed cost identity
        ExtendedPrice = CAST(l.OrderQty * l.UnitPrice AS DECIMAL(16,2)),
        LandedCost    = CAST(l.OrderQty * l.UnitPrice + l.FreightAmount + l.ExpediteFee AS DECIMAL(16,2)),

        -- contracted position at the order date
        cp.AgreementNo,
        ContractedUnitPrice = cp.AgreedUnitPrice,
        AgreementsInForce   = ISNULL(cp.AgreementsInForce, 0),
        IsOnContract        = CAST(CASE WHEN cp.AgreedUnitPrice IS NOT NULL THEN 1 ELSE 0 END AS BIT),
        -- purchase price variance: only meaningful where a contract existed.
        -- Off-contract lines are NOT scored zero variance -- that would report
        -- perfect compliance for the spend nobody negotiated.
        PPVAmount = CASE WHEN cp.AgreedUnitPrice IS NULL THEN NULL
                         ELSE CAST((l.UnitPrice - cp.AgreedUnitPrice) * l.OrderQty AS DECIMAL(16,2)) END,
        ContractedExtended = CASE WHEN cp.AgreedUnitPrice IS NULL THEN NULL
                         ELSE CAST(cp.AgreedUnitPrice * l.OrderQty AS DECIMAL(16,2)) END,

        -- what actually arrived and what could be used
        r.QtyReceived, r.QtyAccepted, r.QtyRejected,
        Receipts = ISNULL(r.Receipts, 0),
        IsReceived = CAST(CASE WHEN r.POLineKey IS NULL THEN 0 ELSE 1 END AS BIT),
        AcceptanceRatePct = CASE WHEN ISNULL(r.QtyReceived,0) = 0 THEN NULL
                                 ELSE CAST(100.0 * r.QtyAccepted / r.QtyReceived AS DECIMAL(6,2)) END,
        -- the figure that changes supplier rankings
        CostPerAcceptedUnit = CASE WHEN ISNULL(r.QtyAccepted,0) = 0 THEN NULL
                                   ELSE CAST((l.OrderQty * l.UnitPrice + l.FreightAmount + l.ExpediteFee)
                                             / r.QtyAccepted AS DECIMAL(16,4)) END,
        -- the cost of what was paid for and cannot be used
        RejectedValue = CASE WHEN ISNULL(r.QtyRejected,0) = 0 THEN 0
                             ELSE CAST(r.QtyRejected * l.UnitPrice AS DECIMAL(16,2)) END,

        -- delivery against the promise the supplier made
        LastReceiptDate = dr.[Date],
        DaysLate = CASE WHEN r.LastReceiptKey IS NULL THEN NULL
                        ELSE DATEDIFF(DAY, dp.[Date], dr.[Date]) END,
        IsOnTime = CASE WHEN r.LastReceiptKey IS NULL THEN NULL
                        ELSE CAST(CASE WHEN r.LastReceiptKey <= l.PromisedDateKey THEN 1 ELSE 0 END AS BIT) END,
        AsOfDate = @AsOf
    FROM dbo.Fact_PurchaseOrderLine l
    JOIN dbo.Dim_Date do ON do.DateKey = l.OrderDateKey
    JOIN dbo.Dim_Date dp ON dp.DateKey = l.PromisedDateKey
    CROSS JOIN AsOfKey a
    LEFT JOIN Rcpt r ON r.POLineKey = l.POLineKey
    LEFT JOIN dbo.Dim_Date dr ON dr.DateKey = r.LastReceiptKey
    OUTER APPLY dbo.fn_ContractedPrice(l.PartKey, l.VendorKey, l.OrderDateKey) cp
    WHERE l.OrderDateKey <= a.k          -- an order placed after @AsOf does not exist yet
);
GO

/*
--------------------------------------------------------------------------------
vw_POLineCost -- the reporting-date snapshot (2025-12-31), enriched with the
part, vendor and buyer attributes every consumer needs. This is the one extract
Power Query and Power BI import.
--------------------------------------------------------------------------------
*/
CREATE VIEW dbo.vw_POLineCost AS
SELECT
    c.POLineKey, c.PONumber, c.POLineNo,
    p.PartNumber, p.PartName, p.Category, p.Criticality,
    p.QualifiedSupplierCount, p.QualificationMonths,
    IsSingleSource = CAST(CASE WHEN p.QualifiedSupplierCount = 1 THEN 1 ELSE 0 END AS BIT),
    v.VendorID, v.VendorName, v.Country, v.VendorTier, v.LumenRevenueSharePct,
    b.BuyerID, b.BuyerName, b.Team,
    c.OrderDate, c.PromisedDate, c.OrderYear, c.OrderYearMonth,
    c.OrderQty, c.UnitPrice, c.FreightAmount, c.ExpediteFee, c.IsExpedited,
    c.ExtendedPrice, c.LandedCost,
    c.AgreementNo, c.ContractedUnitPrice, c.AgreementsInForce, c.IsOnContract,
    c.PPVAmount, c.ContractedExtended,
    c.QtyReceived, c.QtyAccepted, c.QtyRejected, c.Receipts, c.IsReceived,
    c.AcceptanceRatePct, c.CostPerAcceptedUnit, c.RejectedValue,
    c.LastReceiptDate, c.DaysLate, c.IsOnTime,
    c.AsOfDate
FROM dbo.fn_POLineCost('2025-12-31') c
JOIN dbo.Dim_Part   p ON p.PartKey   = c.PartKey
JOIN dbo.Dim_Vendor v ON v.VendorKey = c.VendorKey
JOIN dbo.Dim_Buyer  b ON b.BuyerKey  = c.BuyerKey;
GO

DECLARE @missing VARCHAR(400) = '';
IF OBJECT_ID('dbo.fn_ContractedPrice', 'IF') IS NULL SET @missing += 'fn_ContractedPrice ';
IF OBJECT_ID('dbo.fn_POLineCost', 'IF')      IS NULL SET @missing += 'fn_POLineCost ';
IF OBJECT_ID('dbo.vw_POLineCost', 'V')       IS NULL SET @missing += 'vw_POLineCost ';
IF @missing <> '' THROW 51003, 'FAILED to create core objects -- scroll up for the compile error.', 1;
PRINT 'Core objects created and verified: fn_ContractedPrice, fn_POLineCost(@AsOf), vw_POLineCost.';
GO
