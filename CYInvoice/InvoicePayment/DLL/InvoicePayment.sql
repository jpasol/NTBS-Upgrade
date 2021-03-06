if exists (select * from sysobjects where id = object_id(N'[dbo].[vueTAdjust]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vueTAdjust]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[vueVAdjust]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vueVAdjust]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[vueORList]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vueORList]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[vueInvAdjustment]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vueInvAdjustment]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[vueORList_Adjustment]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vueORList_Adjustment]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[VueFAdjust]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[VueFAdjust]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[viewFinalORList]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[viewFinalORList]
GO


SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

CREATE VIEW dbo.vueTAdjust
AS
SELECT *
FROM INVPAYDTL
WHERE (ORNum IN
        (SELECT ORNUM
      FROM invpayHDR
      WHERE ORTYPE = 'ADJ'))

GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO


SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

CREATE VIEW dbo.vueVAdjust
AS
SELECT INVNum, SUM(PAYAmt) AS Adjustment, remarks
FROM vueTAdjust
GROUP BY INVNum, remarks

GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO


SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

CREATE VIEW dbo.vueORList
AS
SELECT DISTINCT 
    INVPAYHDR.cuscde, INVPAYDTL.ORNum, 
    INVPAYHDR.TotalAMT, INVPAYDTL.INVNum, 
    INVPAYDTL.INVAmt, INVICT.invvat, INVICT.invtax, 
    INVICT.totalpay, INVPAYDTL.PAYAmt, INVPAYDTL.RBalance, 
    INVPAYHDR.AvailAMT, INVPAYHDR.userid
FROM INVPAYDTL INNER JOIN
    INVPAYHDR ON 
    INVPAYDTL.ORNum = INVPAYHDR.ORNum INNER JOIN
    INVICT ON INVPAYDTL.INVNum = INVICT.invnum
WHERE (INVPAYHDR.ORType <> 'ADJ')

GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

CREATE VIEW dbo.vueInvAdjustment
AS
SELECT DISTINCT 
    INVPAYHDR.cuscde, INVPAYDTL.INVNum, 
    SUM(INVPAYDTL.PAYAmt) AS TAdjustment, 
    MAX(DISTINCT INVPAYDTL.remarks) AS Remarks, 
    INVPAYDTL.PAYDate
FROM INVPAYDTL INNER JOIN
    INVPAYHDR ON 
    INVPAYDTL.ORNum = INVPAYHDR.ORNum
WHERE (INVPAYDTL.ORNum IN
        (SELECT ornum
      FROM invpayhdr
      WHERE ortype = 'ADJ'))
GROUP BY INVPAYDTL.INVNum, INVPAYHDR.cuscde, 
    INVPAYDTL.PAYDate

GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO



SET QUOTED_IDENTIFIER  ON    SET ANSI_NULLS  ON 
GO

CREATE VIEW dbo.vueORList_Adjustment
AS
SELECT O.ORNum, O.INVNum, O.INVAmt, O.PAYAmt, I.TAdjustment, 
    I.Remarks, O.RBalance
FROM vueORList O LEFT OUTER JOIN
    vueInvAdjustment I ON O.INVNum = I.invnum

GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO


SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

CREATE VIEW dbo.VueFAdjust
AS
SELECT I.ORNum, I.INVNum, I.INVAmt, I.PAYAmt, V.Adjustment, 
    I.PAYDate, INVPAYHDR.cuscde, INVPAYHDR.CheckNo1, 
    INVPAYHDR.CheckNo2, INVPAYHDR.TotalAMT, 
    INVPAYHDR.ORDate, INVPAYHDR.ORType, V.remarks, 
    INVPAYHDR.AvailAMT
FROM INVPAYDTL I INNER JOIN
    INVPAYHDR ON 
    I.ORNum = INVPAYHDR.ORNum LEFT OUTER JOIN
    vueVAdjust V ON I.INVNum = V.INVNum

GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

CREATE VIEW dbo.viewFinalORList
AS
SELECT DISTINCT 
    vueORList_Adjustment.ORNum, INVPAYHDR.TotalAMT, 
    vueORList_Adjustment.INVNum, 
    vueORList_Adjustment.INVAmt, 
    vueORList_Adjustment.PAYAmt, 
    vueORList_Adjustment.RBalance, 
    vueORList_Adjustment.TAdjustment, INVICT.totalpay, 
    INVPAYHDR.cuscde, INVPAYHDR.CheckNo1, 
    INVPAYHDR.CheckNo2, INVPAYHDR.ORDate, 
    vueORList_Adjustment.Remarks, INVICT.invvat, 
    INVICT.invtax
FROM vueORList_Adjustment INNER JOIN
    INVICT ON 
    vueORList_Adjustment.INVNum = INVICT.invnum LEFT OUTER JOIN
    INVPAYHDR ON 
    vueORList_Adjustment.ORNum = INVPAYHDR.ORNum

GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

Update ornum set orno=10000



