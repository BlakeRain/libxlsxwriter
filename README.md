# Create Excel files with Haskell 

  This repo provides Haskell bindings for C library [libxlsxwriter](http://libxlsxwriter.github.io/)
  
## Example usage

This example uses Yesod web framework 

```
getExcelExportR :: Handler ()
getExcelExportR = do
  lfilter   <- parseLeadFilter Nothing
  data     <- runDB $ filterLeadList lfilter { lfLimit  = Just 1000
                                              , lfOffset = Nothing
                                              }
  gathering <- runDB $ gather [] data

  payload <- liftIO $ do
    path    <- (</>) <$> getTemporaryDirectory <*> replicateM 10 (randomRIO ('a', 'z'))
    wbook   <- workbookNew path
    wsheet  <- workbookAddWorksheet wbook "Excel example"

    titleFormat <- workbookAddFormat wbook
    formatSetBold titleFormat
    formatSetAlign titleFormat AlignCenter

    timeFormat <- workbookAddFormat wbook
    formatSetNumFormat timeFormat "DD/MM/YY HH:MM"

    let titles = [ ("Id",        10)
                 , ("Name", 30)
                 , ("Address",   60)
                 , ("Added",     20)
                 , ("Added By",  20)
                 , ("No. Sales", 10)
                 , ("No. Calls", 10)
                 ]

    forM_ (zip [0 ..] titles) $ \(index, (title, width)) -> do
      worksheetSetColumn wsheet index index width Nothing
      worksheetWriteString wsheet 0 index title (Just titleFormat)

    forM_ (zip [1 ..] data) $ \(row, Entity key item) -> do
      worksheetWriteString wsheet row 0 (unpack (toPathPiece key)) Nothing
      worksheetWriteString wsheet row 1 (unpack (dataListItemName item)) Nothing
      worksheetWriteString wsheet row 2 (unpack (dataListItemAddress item)) Nothing
      worksheetWriteUTCTime wsheet row 3 (dataListItemAdded item) (Just timeFormat)
      worksheetWriteString wsheet row 4 (dataListItemAddedBy item) Nothing 
      worksheetWriteNumber wsheet row 5 (fromRational $ toRational $ dataListItemNumSales item) Nothing
      worksheetWriteNumber wsheet row 6 (fromRational $ toRational $ dataListItemNumCalls item) Nothing

    workbookClose wbook
    blob <- BS.readFile path
    removeFile path
    return blob

  time <- formatTime defaultTimeLocale "%Y-%m-%d" <$> liftIO getCurrentTime
  addHeader "Content-Disposition" [st|attachment; filename="Excel-example #{time}.xlsx"|]
  sendResponse ("application/vnd.ms-excel" :: ContentType, toContent payload)
```

