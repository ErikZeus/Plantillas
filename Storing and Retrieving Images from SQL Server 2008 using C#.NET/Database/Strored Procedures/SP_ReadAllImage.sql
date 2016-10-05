

/****** Object:  StoredProcedure [dbo].[ReadAllImage]    Script Date: 03/26/2012 21:43:43 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


create proc [dbo].[ReadAllImage] as
SELECT * FROM ImageData
GO


