/****** Object:  StoredProcedure [dbo].[ReadImage]    Script Date: 03/26/2012 21:44:51 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


create proc [dbo].[ReadImage] @imgId int as
SELECT ImageData FROM ImageData
WHERE ImageID=@imgId
GO


