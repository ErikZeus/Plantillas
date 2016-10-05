

/****** Object:  StoredProcedure [dbo].[SaveImage]    Script Date: 03/26/2012 21:45:23 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

create proc [dbo].[SaveImage] @img image as
INSERT INTO ImageData(ImageData)
VALUES (@img)
GO


