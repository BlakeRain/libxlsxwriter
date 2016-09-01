
import Control.Monad
import Data.Time
import Data.Default
import Data.Excel

----------------------------------------------------------------------------------------------------

main :: IO ()
main = do
  wb <- workbookNew "test1.xlsx"
  let props = def { docPropertiesTitle   = "Test Workbook"
                  , docPropertiesCompany = "Inchora"
                  }
  workbookSetProperties wb props
  ws <- workbookAddWorksheet wb "First Sheet"
  df <- workbookAddFormat wb
  formatSetNumFormat df "mmm d yyyy hh:mm AM/PM"
  forM_ [0 .. 10] $ \n ->
    worksheetWriteNumber ws n 0 (100.0 * fromIntegral n) Nothing
  now <- getZonedTime
  worksheetSetColumn ws 1 1 20 Nothing
  worksheetWriteDateTime ws 0 1 (zonedTimeToDateTime now) (Just df)
  let io = ImageOptions 0 0 0.5 0.5
  worksheetInsertImageOpt ws 0 2
    "test/image.png" io
  workbookClose wb
