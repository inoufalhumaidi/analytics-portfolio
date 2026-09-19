/*
================================================================================
Project 3 -- Lumen Optics Manufacturing: Photonics Spend Scorecard
Script:  07_uat_test_cases.sql
Purpose: User Acceptance Test suite. Run after any change to the generator, the
         cost layer, the erosion model or the queue.

WHAT A TEST IS FOR HERE
    Not "does the query run". Every case below asserts a property that, if it
    broke, would let the reports keep producing confident, plausible, wrong
    numbers. Several exist because exactly that happened during the build, and
    the comment on each says so.

DESIGN
    Every test writes one row to #UATResults with an expected value, an actual
    value and a verdict, so a failure tells you what it found rather than only
    that something was wrong. The script is read-only and safe to run against
    the reporting database at any time.

    Monetary comparisons use a half-cent tolerance. Exact float equality on
    aggregated DECIMAL arithmetic fails for reasons unrelated to the thing
    being tested.

DATA DISCLOSURE: Lumen Optics Manufacturing is fictional; all data is synthetic.
================================================================================
*/

USE LumenSpend;
GO
SET NOCOUNT ON;
GO

IF OBJECT_ID('tempdb..#UATResults') IS NOT NULL DROP TABLE #UATResults;
CREATE TABLE #UATResults (
    TestID      VARCHAR(10)   NOT NULL,
    Area        VARCHAR(30)   NOT NULL,
    Requirement VARCHAR(200)  NOT NULL,
    Expected    VARCHAR(70)   NOT NULL,
    Actual      VARCHAR(70)   NULL,
    Verdict     VARCHAR(6)    NULL,
    WhyItMatters VARCHAR(400) NOT NULL
);
GO

/* ---------------------------------------------------------------------------
UAT-01  The landed cost identity holds for every purchase order line.
--------------------------------------------------------------------------- */
DECLARE @drift INT = (
    SELECT COUNT(*) FROM dbo.fn_POLineCost('2025-12-31')
    WHERE ABS(LandedCost - (OrderQty * UnitPrice + FreightAmount + ExpediteFee)) > 0.005);
INSERT INTO #UATResults VALUES ('UAT-01','Landed cost',
 'LandedCost equals extended price plus freight plus expedite, for every line',
 '0', CAST(@drift AS VARCHAR(20)), CASE WHEN @drift = 0 THEN 'PASS' ELSE 'FAIL' END,
 'This is the one definition every downstream number depends on. If it drifts for even one line, the spend totals, the supplier rankings and the opportunity are all wrong and nothing downstream can detect it.');

/* ---------------------------------------------------------------------------
UAT-02  As-of integrity: orders and receipts after the reporting date are
        excluded, and the exclusion is on the FACT row, never on a joined date.
--------------------------------------------------------------------------- */
DECLARE @junLines INT = (SELECT COUNT(*) FROM dbo.fn_POLineCost('2025-06-30'));
DECLARE @junDirect INT = (
    SELECT COUNT(*) FROM dbo.Fact_PurchaseOrderLine WHERE OrderDateKey <= 20250630);
DECLARE @junQty BIGINT = (SELECT ISNULL(SUM(QtyReceived),0) FROM dbo.fn_POLineCost('2025-06-30'));
DECLARE @junQtyDirect BIGINT = (
    SELECT ISNULL(SUM(g.QtyReceived),0) FROM dbo.Fact_GoodsReceipt g
    JOIN dbo.Fact_PurchaseOrderLine l ON l.POLineKey = g.POLineKey
    WHERE g.ReceiptDateKey <= 20250630 AND l.OrderDateKey <= 20250630);
INSERT INTO #UATResults VALUES ('UAT-02','As-of integrity',
 'The mid-year snapshot counts only orders placed and receipts booked by then',
 CAST(@junDirect AS VARCHAR(20)) + ' lines / ' + CAST(@junQtyDirect AS VARCHAR(20)) + ' units',
 CAST(@junLines AS VARCHAR(20)) + ' lines / ' + CAST(@junQty AS VARCHAR(20)) + ' units',
 CASE WHEN @junLines = @junDirect AND @junQty = @junQtyDirect THEN 'PASS' ELSE 'FAIL' END,
 'Without this, every historical spend report improves whenever a later receipt lands, and quarter-on-quarter comparisons become meaningless.');

/* ---------------------------------------------------------------------------
UAT-03  An open-ended agreement is still in force.

        Excluding NULL ValidTo is the usual way contract coverage is made to
        look far worse than it is, and it would inflate maverick spend with
        purchases that are perfectly well governed.
--------------------------------------------------------------------------- */
DECLARE @openEnded INT = (SELECT COUNT(*) FROM dbo.Fact_PriceAgreement WHERE ValidToDateKey IS NULL);
DECLARE @coveredByOpen INT = (
    SELECT COUNT(*) FROM dbo.fn_POLineCost('2025-12-31') c
    WHERE c.IsOnContract = 1
      AND EXISTS (SELECT 1 FROM dbo.Fact_PriceAgreement pa
                  WHERE pa.AgreementNo = c.AgreementNo AND pa.ValidToDateKey IS NULL));
INSERT INTO #UATResults VALUES ('UAT-03','Contract resolution',
 'Lines are matched to open-ended agreements, not dropped as uncovered',
 '> 0 lines matched to one of ' + CAST(@openEnded AS VARCHAR(10)) + ' open agreements',
 CAST(@coveredByOpen AS VARCHAR(20)) + ' lines',
 CASE WHEN @openEnded > 0 AND @coveredByOpen > 0 THEN 'PASS' ELSE 'FAIL' END,
 'An open ValidTo means the contract has no end date, not that it never applied. Treating NULL as "not in force" would report perfectly governed spend as maverick.');

/* ---------------------------------------------------------------------------
UAT-04  Off-contract lines carry NULL variance, never zero.
--------------------------------------------------------------------------- */
DECLARE @zeroPPVOffContract INT = (
    SELECT COUNT(*) FROM dbo.fn_POLineCost('2025-12-31')
    WHERE IsOnContract = 0 AND PPVAmount IS NOT NULL);
INSERT INTO #UATResults VALUES ('UAT-04','Price variance',
 'A line with no agreement has no price variance, rather than a variance of zero',
 '0', CAST(@zeroPPVOffContract AS VARCHAR(20)),
 CASE WHEN @zeroPPVOffContract = 0 THEN 'PASS' ELSE 'FAIL' END,
 'Scoring off-contract spend as zero variance reports perfect price compliance for exactly the spend nobody negotiated, which is the most flattering possible lie a procurement report can tell.');

/* ---------------------------------------------------------------------------
UAT-05  The expected-price curve is exactly what the benchmark implies.
--------------------------------------------------------------------------- */
DECLARE @curveDrift DECIMAL(18,6) = (
    SELECT ISNULL(MAX(ABS(ExpectedPrice
         - FirstPrice * POWER(1.0 - BenchmarkErosionPct/100.0, YearsElapsed))), 0)
    FROM dbo.fn_PriceErosion('2025-12-31', 10, 500));
INSERT INTO #UATResults VALUES ('UAT-05','Erosion model',
 'ExpectedPrice equals FirstPrice x (1 - benchmark) ^ years, for every pair',
 '< 0.01', CAST(@curveDrift AS VARCHAR(30)),
 CASE WHEN @curveDrift < 0.01 THEN 'PASS' ELSE 'FAIL' END,
 'The entire opportunity figure is the distance between the paid price and this curve. If the curve is not exactly what the benchmark says, the opportunity is an opinion.');

/* ---------------------------------------------------------------------------
UAT-06  The erosion window cannot overlap itself.

        Caught during the build: the first-90 and last-90 day price windows
        overlap when a pair's history is shorter than 180 days, averaging the
        same orders into both the baseline and the current price and producing
        an erosion rate near zero for any pair. The guard now lives inside the
        function rather than in the caller's arguments.
--------------------------------------------------------------------------- */
DECLARE @shortSpan INT = (
    SELECT COUNT(*) FROM dbo.fn_PriceErosion('2025-12-31', 1, 1)
    WHERE DATEDIFF(DAY, FirstDate, LastDate) < 180);
INSERT INTO #UATResults VALUES ('UAT-06','Erosion model',
 'No pair with under 180 days of history is returned, even at the loosest arguments',
 '0', CAST(@shortSpan AS VARCHAR(20)),
 CASE WHEN @shortSpan = 0 THEN 'PASS' ELSE 'FAIL' END,
 'A guard that lives in the caller''s parameters is not a guard. The next caller passes different numbers and gets a confident erosion rate computed from overlapping windows.');

/* ---------------------------------------------------------------------------
UAT-07  Erosion capture is consistent with its own components.
--------------------------------------------------------------------------- */
DECLARE @captureDrift DECIMAL(18,4) = (
    SELECT ISNULL(MAX(ABS(ErosionCapturePct - 100.0 * ActualErosionPct / NULLIF(BenchmarkErosionPct,0))), 0)
    FROM dbo.fn_PriceErosion('2025-12-31', 10, 500) WHERE BenchmarkErosionPct > 0);
INSERT INTO #UATResults VALUES ('UAT-07','Erosion model',
 'ErosionCapturePct equals actual erosion divided by the benchmark',
 '< 0.01', CAST(@captureDrift AS VARCHAR(30)),
 CASE WHEN @captureDrift < 0.01 THEN 'PASS' ELSE 'FAIL' END,
 'The headline KPI must agree with the two numbers printed beside it, or a reader doing the division on the page gets a different answer than the report.');

/* ---------------------------------------------------------------------------
UAT-08  Every queued pair carries exactly one known action and an instruction.
--------------------------------------------------------------------------- */
DECLARE @badAction INT = (
    SELECT COUNT(*) FROM dbo.fn_RenegotiationQueue('2025-12-31', 8)
    WHERE ActionCode NOT IN ('RENEGOTIATE','PUT_ON_CONTRACT','DUAL_SOURCE','QUALIFY_ALTERNATE','FIX_QUALITY','ACCEPT')
       OR ActionCode IS NULL OR RecommendedAction IS NULL);
INSERT INTO #UATResults VALUES ('UAT-08','Renegotiation queue',
 'Every pair carries one recognised action code and a written instruction',
 '0', CAST(@badAction AS VARCHAR(20)),
 CASE WHEN @badAction = 0 THEN 'PASS' ELSE 'FAIL' END,
 'A CASE ladder that falls through leaves a category manager with a row and no instruction. The queue stops being a control the moment a row does not say what to do.');

/* ---------------------------------------------------------------------------
UAT-09  Nothing material is parked as ACCEPT.

        Caught during the build: a $179k annual opportunity was classified
        ACCEPT because its leverage score was low. Weak leverage is a reason to
        expect a hard conversation, not a reason to skip one.
--------------------------------------------------------------------------- */
DECLARE @bigAccept DECIMAL(16,2) = (
    SELECT ISNULL(MAX(AnnualOpportunity), 0) FROM dbo.fn_RenegotiationQueue('2025-12-31', 8)
    WHERE ActionCode = 'ACCEPT');
INSERT INTO #UATResults VALUES ('UAT-09','Renegotiation queue',
 'No opportunity of $100,000 or more is classified ACCEPT',
 '< 100000', CAST(@bigAccept AS VARCHAR(30)),
 CASE WHEN @bigAccept < 100000 THEN 'PASS' ELSE 'FAIL' END,
 'Parking six figures a year as "not worth pursuing" because the supplier is large is how a sourcing report quietly gives money away. Above a threshold you ask regardless of leverage.');

/* ---------------------------------------------------------------------------
UAT-10  The quarter's worklist respects category-manager capacity.
--------------------------------------------------------------------------- */
DECLARE @overCapacity INT = (
    SELECT COUNT(*) FROM (
        SELECT BuyerID, Extra = SUM(CASE WHEN IsThisQuarter = 1 AND BuyerRank > 8 THEN 1 ELSE 0 END)
        FROM dbo.fn_RenegotiationQueue('2025-12-31', 8) GROUP BY BuyerID) x
    WHERE x.Extra > 0);
INSERT INTO #UATResults VALUES ('UAT-10','Renegotiation queue',
 'No category manager is given more than eight negotiations in the quarter',
 '0 buyers', CAST(@overCapacity AS VARCHAR(20)),
 CASE WHEN @overCapacity = 0 THEN 'PASS' ELSE 'FAIL' END,
 'A list of 185 negotiations is not a plan. If the capacity rule leaks, the queue becomes a ranked ledger and the sourcing team goes back to working on whatever is loudest.');

/* ---------------------------------------------------------------------------
UAT-11  The negotiation is assigned to the buyer who actually places the spend.

        Caught during the build: ownership was assigned with MIN(BuyerKey), and
        140 of 185 pairs are touched by more than one buyer -- so for three
        quarters of the queue the assignment was an accident of key ordering.
--------------------------------------------------------------------------- */
DECLARE @wrongBuyer INT = (
    SELECT COUNT(*) FROM dbo.fn_RenegotiationQueue('2025-12-31', 8) q
    WHERE q.BuyerID <> (
        SELECT TOP 1 b.BuyerID
        FROM dbo.fn_POLineCost('2025-12-31') c
        JOIN dbo.Dim_Buyer b ON b.BuyerKey = c.BuyerKey
        JOIN dbo.Dim_Part  p ON p.PartKey  = c.PartKey
        JOIN dbo.Dim_Vendor v ON v.VendorKey = c.VendorKey
        WHERE p.PartNumber = q.PartNumber AND v.VendorID = q.VendorID
          AND c.OrderDate > DATEADD(MONTH, -12, '2025-12-31')
        GROUP BY b.BuyerID, b.BuyerKey
        ORDER BY SUM(c.ExtendedPrice) DESC, b.BuyerKey));
INSERT INTO #UATResults VALUES ('UAT-11','Renegotiation queue',
 'The assigned buyer is the one who placed the most spend on that pair',
 '0 mismatches', CAST(@wrongBuyer AS VARCHAR(20)),
 CASE WHEN @wrongBuyer = 0 THEN 'PASS' ELSE 'FAIL' END,
 'Most pairs are bought by more than one person. Assigning the negotiation by key order hands it to someone who may never have placed an order for the part.');

/* ---------------------------------------------------------------------------
UAT-12  The cost index equals the like-for-like index, recomputed at PAIR grain.

        The earlier form of this test asserted only "index >= 100". That is
        true under ANY positive weighting of the pairs, so it passed happily
        while fn_VendorScorecard was summing each pair's landed cost once per
        PO line -- a join fan-out that turned the index into a line-count-
        weighted average of itself and moved 33 of 40 vendors' rank positions.
        A test that cannot fail for the reason it was written is not a test.
        This version recomputes the index at the grain the definition is stated
        in, which is the grain the implementation got wrong.
--------------------------------------------------------------------------- */
DECLARE @indexUnder100 INT = (SELECT COUNT(*) FROM dbo.fn_VendorScorecard('2025-12-31') WHERE CostIndexVsBest < 99.99);
DECLARE @indexMismatch INT;
;WITH C AS (SELECT * FROM dbo.fn_POLineCost('2025-12-31') WHERE PONumber NOT LIKE 'PO-D%'),
PV AS (SELECT PartKey, VendorKey, Landed = SUM(LandedCost), Accepted = SUM(ISNULL(QtyAccepted,0)),
              CPA = SUM(LandedCost) / NULLIF(SUM(ISNULL(QtyAccepted,0)), 0)
       FROM C GROUP BY PartKey, VendorKey HAVING SUM(ISNULL(QtyAccepted,0)) > 0),
PB AS (SELECT PartKey, BestCPA = MIN(CPA) FROM PV GROUP BY PartKey),
Expect AS (SELECT pv.VendorKey,
                  Idx = CAST(100.0 * SUM(pv.Landed) / NULLIF(SUM(pb.BestCPA * pv.Accepted), 0) AS DECIMAL(9,2))
           FROM PV pv JOIN PB pb ON pb.PartKey = pv.PartKey
           GROUP BY pv.VendorKey)
SELECT @indexMismatch = COUNT(*)
FROM dbo.fn_VendorScorecard('2025-12-31') s
JOIN Expect e ON e.VendorKey = s.VendorKey
WHERE ABS(s.CostIndexVsBest - e.Idx) > 0.01;

INSERT INTO #UATResults VALUES ('UAT-12','Vendor scorecard',
 'CostIndexVsBest is at least 100 and equals the index recomputed at pair grain',
 '0 under 100, 0 mismatched',
 CAST(@indexUnder100 AS VARCHAR(20)) + ' / ' + CAST(@indexMismatch AS VARCHAR(20)),
 CASE WHEN @indexUnder100 = 0 AND @indexMismatch = 0 THEN 'PASS' ELSE 'FAIL' END,
 'The index compares a vendor with the cheapest cost per accepted unit on the same parts. Below 100 would mean beating the best, which is arithmetically impossible. Equality with the pair-grain recomputation is the part that catches a fan-out, which leaves the value above 100 and entirely plausible.');

/* ---------------------------------------------------------------------------
UAT-13  The vendor scorecard ties back to line level.
--------------------------------------------------------------------------- */
DECLARE @vendorSpend DECIMAL(18,2) = (SELECT SUM(TotalSpend) FROM dbo.fn_VendorScorecard('2025-12-31'));
-- The scorecard reports COMMITTED spend -- every non-duplicate line, whether
-- delivered yet or not -- while acceptance and the cost index are computed only
-- over what has actually arrived. The earlier form of this test compared
-- against received lines only and failed on its own premise, not on the code.
DECLARE @lineSpend DECIMAL(18,2) = (
    SELECT SUM(c.ExtendedPrice) FROM dbo.fn_POLineCost('2025-12-31') c
    WHERE c.PONumber NOT LIKE 'PO-D%');
INSERT INTO #UATResults VALUES ('UAT-13','Aggregation',
 'Vendor-level spend equals line-level committed spend across non-duplicate lines',
 CAST(@lineSpend AS VARCHAR(40)), CAST(@vendorSpend AS VARCHAR(40)),
 CASE WHEN ABS(@vendorSpend - @lineSpend) < 0.05 THEN 'PASS' ELSE 'FAIL' END,
 'A join that fans out, or a vendor with no scorecard row, quietly changes the total. This is the cheapest possible check that the rollup did not invent or lose money.');

/* ---------------------------------------------------------------------------
UAT-14  Procedures reject bad input rather than answering a different question.
--------------------------------------------------------------------------- */
DECLARE @g1 VARCHAR(6) = 'FAIL', @g2 VARCHAR(6) = 'FAIL';
BEGIN TRY EXEC dbo.usp_CategorySummary @GroupBy = 'DROP TABLE'; END TRY BEGIN CATCH SET @g1 = 'PASS'; END CATCH
BEGIN TRY EXEC dbo.usp_PartPriceHistory @PartNumber = 'LOM-9999'; END TRY BEGIN CATCH SET @g2 = 'PASS'; END CATCH
INSERT INTO #UATResults VALUES ('UAT-14','Interface',
 'usp_CategorySummary and usp_PartPriceHistory raise on invalid arguments',
 'PASS / PASS', @g1 + ' / ' + @g2,
 CASE WHEN @g1 = 'PASS' AND @g2 = 'PASS' THEN 'PASS' ELSE 'FAIL' END,
 'A procedure that silently falls back to a default answers a question nobody asked, and the caller cannot tell. Raising is the only honest response to an argument it does not understand.');

/* ---------------------------------------------------------------------------
UAT-15  The duplicate detector fires on the planted defects, by behaviour.
--------------------------------------------------------------------------- */
DECLARE @dupes INT = (SELECT Anomalies FROM dbo.vw_DQ_Summary WHERE AnomalyType = 'DUPLICATE_PO_LINE');
DECLARE @planted INT = (SELECT COUNT(*) FROM dbo.Fact_PurchaseOrderLine WHERE PONumber LIKE 'PO-D%');
INSERT INTO #UATResults VALUES ('UAT-15','Data quality',
 'The duplicate check finds every planted duplicate without reading its prefix',
 CAST(@planted AS VARCHAR(10)), CAST(@dupes AS VARCHAR(10)),
 CASE WHEN @dupes = @planted AND @planted > 0 THEN 'PASS' ELSE 'FAIL' END,
 'The quality layer is only worth anything if it is known to fire. Planting a defect of known size and asserting the detector finds exactly that many is the only way to tell a working control from a silent one.');

/* ---------------------------------------------------------------------------
UAT-16  REGRESSION: one defect is not reported twice.

        Caught during the build: duplicated PO lines are inserted after receipts
        are generated, so they can never have one -- and every duplicate was
        also being reported as NEVER_RECEIVED, padding the defect count with a
        consequence of something already on the report.
--------------------------------------------------------------------------- */
DECLARE @neverRecvDupes INT = (
    SELECT COUNT(*) FROM dbo.vw_DQ_Anomalies a
    WHERE a.AnomalyType = 'NEVER_RECEIVED' AND a.EntityRef LIKE 'PO-D%');
INSERT INTO #UATResults VALUES ('UAT-16','Data quality',
 'No duplicate PO line is also reported as never received',
 '0', CAST(@neverRecvDupes AS VARCHAR(20)),
 CASE WHEN @neverRecvDupes = 0 THEN 'PASS' ELSE 'FAIL' END,
 'Counting the cause and its consequence as two defects inflates the headline and makes a quality report look worse than the data is, which erodes trust in the numbers that are real.');

/* ---------------------------------------------------------------------------
UAT-17  The ambiguous-contract check finds the overlapping agreements.
--------------------------------------------------------------------------- */
DECLARE @ambig INT = (SELECT Anomalies FROM dbo.vw_DQ_Summary WHERE AnomalyType = 'AMBIGUOUS_CONTRACT');
DECLARE @overlaps INT = (SELECT COUNT(*) FROM dbo.Fact_PriceAgreement WHERE AgreementNo LIKE 'PA-X%');
INSERT INTO #UATResults VALUES ('UAT-17','Data quality',
 'Overlapping agreements are detected as ambiguous contracted prices',
 '> 0 lines from ' + CAST(@overlaps AS VARCHAR(10)) + ' overlapping agreements',
 CAST(@ambig AS VARCHAR(20)) + ' lines',
 CASE WHEN @overlaps > 0 AND @ambig > 0 THEN 'PASS' ELSE 'FAIL' END,
 'When two agreements cover the same date, the contracted price -- and therefore the price variance on that line -- depends on which row the query reaches first. Reporting the variance without reporting the ambiguity is a number with no error bar.');

/* ---------------------------------------------------------------------------
UAT-18  Contract resolution is deterministic under overlap.
--------------------------------------------------------------------------- */
DECLARE @nondeterministic INT = (
    SELECT COUNT(*) FROM (
        SELECT c.POLineKey, Runs = COUNT(DISTINCT c.ContractedUnitPrice)
        FROM (SELECT POLineKey, ContractedUnitPrice FROM dbo.fn_POLineCost('2025-12-31')
              UNION ALL
              SELECT POLineKey, ContractedUnitPrice FROM dbo.fn_POLineCost('2025-12-31')) c
        GROUP BY c.POLineKey HAVING COUNT(DISTINCT c.ContractedUnitPrice) > 1) x);
INSERT INTO #UATResults VALUES ('UAT-18','Contract resolution',
 'Two evaluations of the same line return the same contracted price',
 '0', CAST(@nondeterministic AS VARCHAR(20)),
 CASE WHEN @nondeterministic = 0 THEN 'PASS' ELSE 'FAIL' END,
 'Without a full tie-break, two overlapping agreements return whichever row the engine reaches first, and the same report gives different variance figures on different runs.');

/* ---------------------------------------------------------------------------
UAT-19  Acceptance arithmetic holds on every receipt.
--------------------------------------------------------------------------- */
DECLARE @acceptDrift INT = (
    SELECT COUNT(*) FROM dbo.Fact_GoodsReceipt WHERE QtyAccepted + QtyRejected <> QtyReceived);
INSERT INTO #UATResults VALUES ('UAT-19','Quality data',
 'Accepted plus rejected equals received, on every receipt',
 '0', CAST(@acceptDrift AS VARCHAR(20)),
 CASE WHEN @acceptDrift = 0 THEN 'PASS' ELSE 'FAIL' END,
 'A receipt whose parts do not add up makes every yield figure and every cost-per-accepted-unit wrong, and no totals check would notice because the total still balances.');

/* ---------------------------------------------------------------------------
UAT-20  The generator is deterministic.
--------------------------------------------------------------------------- */
DECLARE @poChk INT = (SELECT CHECKSUM_AGG(CHECKSUM(PONumber, OrderQty, UnitPrice, OrderDateKey)) FROM dbo.Fact_PurchaseOrderLine);
DECLARE @grChk INT = (SELECT CHECKSUM_AGG(CHECKSUM(ReceiptNo, QtyReceived, QtyAccepted, ReceiptDateKey)) FROM dbo.Fact_GoodsReceipt);
DECLARE @paChk INT = (SELECT CHECKSUM_AGG(CHECKSUM(AgreementNo, AgreedUnitPrice, ValidFromDateKey)) FROM dbo.Fact_PriceAgreement);
INSERT INTO #UATResults VALUES ('UAT-20','Reproducibility',
 'Regenerating the dataset reproduces the documented fingerprints',
 '-1747321861 / -849044980 / -1705852232',
 CAST(@poChk AS VARCHAR(20)) + ' / ' + CAST(@grChk AS VARCHAR(20)) + ' / ' + CAST(@paChk AS VARCHAR(20)),
 CASE WHEN @poChk = -1747321861 AND @grChk = -849044980 AND @paChk = -1705852232 THEN 'PASS' ELSE 'FAIL' END,
 'Every figure quoted in the case study is only checkable if the dataset can be rebuilt byte for byte. Hash-based pseudo-randomness gives that; RAND() and NEWID() do not.');

/* ---------------------------------------------------------------------------
UAT-21  Opportunity is never negative, and the offset is disclosed separately.
--------------------------------------------------------------------------- */
DECLARE @negInQueue INT = (
    SELECT COUNT(*) FROM dbo.fn_RenegotiationQueue('2025-12-31', 8) WHERE AnnualOpportunity < 0);
DECLARE @aheadPairs INT = (
    SELECT COUNT(*) FROM dbo.fn_PriceErosion('2025-12-31', 10, 500) WHERE AnnualOpportunity < 0);
INSERT INTO #UATResults VALUES ('UAT-21','Erosion model',
 'The queue carries no negative opportunity, while pairs ahead of the curve still exist',
 '0 in queue, > 0 ahead of curve',
 CAST(@negInQueue AS VARCHAR(10)) + ' / ' + CAST(@aheadPairs AS VARCHAR(10)),
 CASE WHEN @negInQueue = 0 AND @aheadPairs > 0 THEN 'PASS' ELSE 'FAIL' END,
 'A supplier already beating the curve is not a savings source, so its negative gap must not be netted against another supplier''s shortfall -- that would understate the actionable target and imply an offset that cannot be realised.');
GO

/* ------------------------------ report ---------------------------------- */
SELECT TestID, Area, Requirement, Expected, Actual, Verdict FROM #UATResults ORDER BY TestID;

DECLARE @pass INT = (SELECT COUNT(*) FROM #UATResults WHERE Verdict = 'PASS');
DECLARE @fail INT = (SELECT COUNT(*) FROM #UATResults WHERE Verdict = 'FAIL');
PRINT '';
PRINT 'UAT RESULT: ' + CAST(@pass AS VARCHAR(10)) + ' passed, ' + CAST(@fail AS VARCHAR(10)) + ' failed.';
IF @fail > 0
BEGIN
    SELECT TestID, Requirement, Expected, Actual, WhyItMatters FROM #UATResults WHERE Verdict = 'FAIL' ORDER BY TestID;
    THROW 51020, 'UAT FAILED -- see the failing cases above. Do not publish figures from this build.', 1;
END
PRINT 'All acceptance criteria met.';
GO
