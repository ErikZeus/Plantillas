

/****** Object:  StoredProcedure [dbo].[ReadAllImageIDs]    Script Date: 03/26/2012 21:44:29 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

create proc [dbo].[ReadAllImageIDs] as
SELECT ImageID FROM ImageData
GO


