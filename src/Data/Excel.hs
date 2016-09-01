{-# LANGUAGE ForeignFunctionInterface #-}

--
-- |
-- Module     : Data.Excel
-- Maintainer : Blake Rain <blake.rain@inchora.com>
--
-- Interface to @libxlsxwriter@.
--

module Data.Excel
       ( Workbook
       , workbookNew
       , workbookNewConstantMem
       , workbookClose
       , workbookAddWorksheet
       , workbookAddFormat
       , workbookDefineName
       , DocProperties (..)
       , workbookSetProperties
       , Worksheet
       , Row
       , Col
       , worksheetWriteNumber
       , worksheetWriteString
       , worksheetWriteFormula
       , worksheetWriteArrayFormula
       , DateTime (..)
       , utcTimeToDateTime
       , zonedTimeToDateTime
       , worksheetWriteDateTime
       , worksheetWriteUrl
       , worksheetSetRow
       , worksheetSetColumn
       , ImageOptions (..)
       , worksheetInsertImage
       , worksheetInsertImageOpt
       , worksheetMergeRange
       , worksheetFreezePanes
       , worksheetSplitPanes
       , worksheetSetLandscape
       , worksheetSetPortrait
       , worksheetSetPageView
       , PaperSize (..)
       , worksheetSetPaperSize
       , worksheetSetMargins
       , worksheetSetHeaderCtl
       , worksheetSetFooterCtl
       , worksheetSetZoom
       , worksheetSetPrintScale
       , Format
       , formatSetFontName
       , formatSetFontSize
       , Color (..)
       , formatSetFontColor
       , formatSetNumFormat
       , formatSetBold
       , formatSetItalic
       , UnderlineStyle (..)
       , formatSetUnderline
       , formatSetStrikeout
       , ScriptStyle (..)
       , formatSetScript
       , formatSetBuiltInFormat
       , Align (..)
       , VerticalAlign (..)
       , formatSetAlign
       , formatSetVerticalAlign
       , formatSetTextWrap
       , formatSetRotation
       , formatSetShrink
       , Pattern (..)
       , formatSetPattern
       , formatSetBackgroundColor
       , formatSetForegroundColor
       , Border (..)
       , BorderStyle (..)
       , formatSetBorder
       , formatSetBorderColor
       ) where

----------------------------------------------------------------------------------------------------

import Data.Default
import Data.Word
import Data.Time
import Data.Time.Clock.POSIX

import Foreign
import Foreign.C.String
import Foreign.C.Types

----------------------------------------------------------------------------------------------------

data LxwWorkbook_
newtype Workbook = Workbook (Ptr LxwWorkbook_)

workbookNew :: FilePath -> IO Workbook
workbookNew path = withCString path $ \cpath ->
  Workbook <$> workbook_new cpath

data WorkbookOptions =
  WorkbookOptions { workbookOptionsConstantMem :: Bool
                  }

instance Storable WorkbookOptions where
  sizeOf _ = sizeOf (undefined :: Word8)
  alignment _ = alignment (undefined :: Int)
  peek ptr = do
    cm <- peekByteOff ptr 0 :: IO Word8
    return (WorkbookOptions (cm /= 1))
  poke ptr opts = do
    let cm | workbookOptionsConstantMem opts = 1
           | otherwise = 0
    pokeByteOff ptr 0 (cm :: Word8)

workbookNewConstantMem :: FilePath -> IO Workbook
workbookNewConstantMem path =
  with (WorkbookOptions True) $ \copts ->
    withCString path $ \cpath -> 
      Workbook <$> workbook_new_opt cpath copts

workbookClose :: Workbook -> IO ()
workbookClose (Workbook wb) =
  workbook_close wb

workbookAddWorksheet :: Workbook -> String -> IO Worksheet
workbookAddWorksheet (Workbook wb) name =
  withCString name $ \cname -> do
    Worksheet <$> workbook_add_worksheet wb cname

workbookAddFormat :: Workbook -> IO Format
workbookAddFormat (Workbook wb) =
  Format <$> workbook_add_format wb

data DocProperties =
  DocProperties { docPropertiesTitle         :: String
                , docPropertiesSubject       :: String
                , docPropertiesAuthor        :: String
                , docPropertiesManager       :: String
                , docPropertiesCompany       :: String
                , docPropertiesCategory      :: String
                , docPropertiesKeywords      :: String
                , docPropertiesComments      :: String
                , docPropertiesStatus        :: String
                , docPropertiesHyperlinkBase :: String
                , docPropertiesCreated       :: UTCTime {-CTime-}
               } 

instance Default DocProperties where
  def = DocProperties { docPropertiesTitle = ""
                      , docPropertiesSubject       = ""
                      , docPropertiesAuthor        = ""
                      , docPropertiesManager       = ""
                      , docPropertiesCompany       = ""
                      , docPropertiesCategory      = ""
                      , docPropertiesKeywords      = ""
                      , docPropertiesComments      = ""
                      , docPropertiesStatus        = ""
                      , docPropertiesHyperlinkBase = ""
                      , docPropertiesCreated       =
                          read "1984-07-06 18:00:00 UTC"
                      }

data DocProperties' =
  DocProperties' { docProps'Title         :: CString
                 , docProps'Subject       :: CString
                 , docProps'Author        :: CString
                 , docProps'Manager       :: CString
                 , docProps'Company       :: CString
                 , docProps'Category      :: CString
                 , docProps'Keywords      :: CString
                 , docProps'Comments      :: CString
                 , docProps'Status        :: CString
                 , docProps'HyperlinkBase :: CString
                 , docProps'Created       :: CTime
                 }

withDocProperties :: DocProperties -> (Ptr DocProperties' -> IO a) -> IO a
withDocProperties props action =
  withCString (docPropertiesTitle props) $ \ctitle ->
  withCString (docPropertiesSubject props) $ \csubject ->
  withCString (docPropertiesAuthor props) $ \cauthor ->
  withCString (docPropertiesManager props) $ \cmanager -> 
  withCString (docPropertiesCompany props) $ \ccompany ->
  withCString (docPropertiesCategory props) $ \ccat ->
  withCString (docPropertiesKeywords props) $ \ckws ->
  withCString (docPropertiesComments props) $ \ccmts ->
  withCString (docPropertiesStatus props) $ \cstat -> 
  withCString (docPropertiesHyperlinkBase props) $ \clb ->
    let time   = CTime (round (utcTimeToPOSIXSeconds (docPropertiesCreated props)))
        props' = DocProperties' ctitle csubject cauthor cmanager
                                ccompany ccat ckws ccmts cstat clb time
    in with props' action

instance Storable DocProperties' where
  sizeOf _ = 10 * sizeOf (undefined :: CString) +
                  sizeOf (undefined :: CTime)
  alignment _ = alignment (undefined :: CString)
  peek = error "No implementation of 'peek' for 'DocProperties'"
  poke ptr props = do
    let n = sizeOf (undefined :: CString)
    pokeByteOff ptr (0 * n) (docProps'Title props)
    pokeByteOff ptr (1 * n) (docProps'Subject props)
    pokeByteOff ptr (2 * n) (docProps'Author props)
    pokeByteOff ptr (3 * n) (docProps'Manager props)
    pokeByteOff ptr (4 * n) (docProps'Company props)
    pokeByteOff ptr (5 * n) (docProps'Category props)
    pokeByteOff ptr (6 * n) (docProps'Keywords props)
    pokeByteOff ptr (7 * n) (docProps'Comments props)
    pokeByteOff ptr (8 * n) (docProps'Status props)
    pokeByteOff ptr (9 * n) (docProps'HyperlinkBase props)
    pokeByteOff ptr (10 * n) (docProps'Created props)

workbookSetProperties :: Workbook -> DocProperties -> IO ()
workbookSetProperties (Workbook wb) props =
  withDocProperties props $ \cprops ->
    workbook_set_properties wb cprops

workbookDefineName :: Workbook -> String -> String -> IO ()
workbookDefineName (Workbook wb) name formula =
  withCString name $ \cname -> withCString formula $ \cformula ->
    workbook_define_name wb cname cformula

----------------------------------------------------------------------------------------------------

data LxwWorksheet_
newtype Worksheet = Worksheet (Ptr LxwWorksheet_)

type Row = Word32
type Col = Word16

worksheetWriteNumber :: Worksheet ->
                        Row -> Col ->
                        Double -> Maybe Format -> IO ()
worksheetWriteNumber (Worksheet ws) row col number mfmt =
  worksheet_write_number ws row col number (maybe nullPtr unFormat mfmt)

worksheetWriteString :: Worksheet ->
                        Row -> Col ->
                        String -> Maybe Format -> IO ()
worksheetWriteString (Worksheet ws) row col str mfmt =
  withCString str $ \cstr ->
    worksheet_write_string ws row col cstr (maybe nullPtr unFormat mfmt)

worksheetWriteFormula :: Worksheet ->
                         Row -> Col ->
                         String -> Maybe Format -> IO ()
worksheetWriteFormula (Worksheet ws) row col str mfmt =
  withCString str $ \cstr ->
    worksheet_write_formula ws row col cstr (maybe nullPtr unFormat mfmt)

worksheetWriteArrayFormula :: Worksheet ->
                              Row -> Col ->
                              Row -> Col ->
                              String -> Maybe Format -> IO ()
worksheetWriteArrayFormula (Worksheet ws) frow fcol erow ecol str mfmt =
  withCString str $ \cstr ->
    worksheet_write_array_formula ws frow fcol erow ecol cstr
      (maybe nullPtr unFormat mfmt)

data DateTime =
  DateTime { dtYear   :: CInt
           , dtMonth  :: CInt
           , dtDay    :: CInt
           , dtHour   :: CInt
           , dtMinute :: CInt
           , dtSecond :: CDouble
           }
  deriving (Show)

instance Storable DateTime where
  sizeOf _ = 5 * sizeOf (undefined :: Int) +
                 sizeOf (undefined :: Double)
  alignment _ = alignment (undefined :: Int)
  peek ptr = do
    let ptr' = castPtr ptr
    DateTime <$> peekElemOff ptr' 0
             <*> peekElemOff ptr' 1
             <*> peekElemOff ptr' 2
             <*> peekElemOff ptr' 3
             <*> peekElemOff ptr' 4
             <*> peekElemOff ptr' 5
  poke ptr (DateTime y m d h mi s) = do
    pokeByteOff ptr 0 y
    pokeByteOff ptr 4 m
    pokeByteOff ptr 8 d
    pokeByteOff ptr 12 h
    pokeByteOff ptr 16 mi
    pokeByteOff ptr 20 s

utcTimeToDateTime :: UTCTime -> DateTime
utcTimeToDateTime (UTCTime day time) =
  let (y, m, d)        = toGregorian day
      TimeOfDay h mi s = timeToTimeOfDay time
  in DateTime (fromIntegral y) (fromIntegral m) (fromIntegral d)
       (fromIntegral h) (fromIntegral mi) (fromRational (toRational s))

zonedTimeToDateTime :: ZonedTime -> DateTime
zonedTimeToDateTime = utcTimeToDateTime . zonedTimeToUTC

worksheetWriteDateTime :: Worksheet ->
                          Row -> Col ->
                          DateTime -> Maybe Format -> IO ()
worksheetWriteDateTime (Worksheet ws) row col dt mfmt =
  with dt $ \pdt -> do
    worksheet_write_datetime ws row col pdt (maybe nullPtr unFormat mfmt)

worksheetWriteUrl :: Worksheet ->
                     Row -> Col ->
                     String -> Maybe Format -> IO ()
worksheetWriteUrl (Worksheet ws) row col str mfmt =
  withCString str $ \cstr ->
    worksheet_write_url ws row col cstr (maybe nullPtr unFormat mfmt)

worksheetSetRow :: Worksheet -> Row -> Double -> Maybe Format -> IO ()
worksheetSetRow (Worksheet ws) row height mfmt =
  worksheet_set_row ws row height (maybe nullPtr unFormat mfmt)

worksheetSetColumn :: Worksheet -> Col -> Col -> Double -> Maybe Format -> IO ()
worksheetSetColumn (Worksheet ws) fcol lcol width mfmt =
  worksheet_set_column ws fcol lcol width (maybe nullPtr unFormat mfmt)

data ImageOptions =
  ImageOptions { imageOffsetX :: Int32
               , imageOffsetY :: Int32
               , imageScaleX  :: Double
               , imageScaleY  :: Double
               }

instance Storable ImageOptions where
  sizeOf _ = 2 * sizeOf (undefined :: Int32) +
             2 * sizeOf (undefined :: Double)
  alignment _ = alignment (undefined :: Int32)
  peek ptr =
    ImageOptions <$> peekByteOff ptr 0
                 <*> peekByteOff ptr 4
                 <*> peekByteOff ptr 8
                 <*> peekByteOff ptr 16
  poke ptr (ImageOptions ox oy sx sy) = do
    pokeByteOff ptr 0 ox
    pokeByteOff ptr 4 oy
    pokeByteOff ptr 8 sx
    pokeByteOff ptr 16 sy    

worksheetInsertImage :: Worksheet -> Row -> Col -> FilePath -> IO ()
worksheetInsertImage (Worksheet ws) row col path =
  withCString path $ \cpath ->
    worksheet_insert_image ws row col cpath

worksheetInsertImageOpt :: Worksheet -> Row -> Col ->
                           FilePath -> ImageOptions -> IO ()
worksheetInsertImageOpt (Worksheet ws) row col path opt =
  withCString path $ \cpath -> with opt $ \optr -> 
    worksheet_insert_image_opt ws row col cpath optr

worksheetMergeRange :: Worksheet ->
                       Row -> Col ->
                       Row -> Col ->
                       String -> Maybe Format -> IO ()
worksheetMergeRange (Worksheet ws) frow fcol lrow lcol str mfmt =
  withCString str $ \cstr ->
    worksheet_merge_range ws frow fcol lrow lcol cstr (maybe nullPtr unFormat mfmt)

worksheetFreezePanes :: Worksheet -> Row -> Col -> IO ()
worksheetFreezePanes (Worksheet ws) row col =
  worksheet_freeze_panes ws row col

worksheetSplitPanes :: Worksheet -> Double -> Double -> IO ()
worksheetSplitPanes (Worksheet ws) vertical horizontal =
  worksheet_split_panes ws vertical horizontal

worksheetSetLandscape :: Worksheet -> IO ()
worksheetSetLandscape (Worksheet ws) =
  worksheet_set_landscape ws

worksheetSetPortrait :: Worksheet -> IO ()
worksheetSetPortrait (Worksheet ws) =
  worksheet_set_portrait ws

worksheetSetPageView :: Worksheet -> IO ()
worksheetSetPageView (Worksheet ws) =
  worksheet_set_page_view ws

data PaperSize
  = DefaultPaper
  | LetterPaper
  | A3Paper
  | A4Paper
  | A5Paper
  | OtherPaper Word8
  deriving (Eq)

worksheetSetPaperSize :: Worksheet -> PaperSize -> IO ()
worksheetSetPaperSize (Worksheet ws) paper =
  worksheet_set_paper ws (toPaper paper)
  where
    toPaper :: PaperSize -> Word8
    toPaper DefaultPaper   = 0
    toPaper LetterPaper    = 1
    toPaper A3Paper        = 8
    toPaper A4Paper        = 9
    toPaper A5Paper        = 11
    toPaper (OtherPaper n) = n

worksheetSetMargins :: Worksheet ->
                       Double -> Double -> Double -> Double -> IO ()
worksheetSetMargins (Worksheet ws) left right top bottom =
  worksheet_set_margins ws left right top bottom

worksheetSetHeaderCtl :: Worksheet -> String -> IO ()
worksheetSetHeaderCtl (Worksheet ws) str =
  withCString str $ \cstr -> worksheet_set_header ws cstr

worksheetSetFooterCtl :: Worksheet -> String -> IO ()
worksheetSetFooterCtl (Worksheet ws) str =
  withCString str $ \cstr -> worksheet_set_footer ws cstr

worksheetSetZoom :: Worksheet -> Double -> IO ()
worksheetSetZoom (Worksheet ws) zoom =
  worksheet_set_zoom ws (round (100.0 * zoom'))
  where
    zoom' = min 0.1 (max 4.0 zoom)

worksheetSetPrintScale :: Worksheet -> Double -> IO ()
worksheetSetPrintScale (Worksheet ws) scale =
  worksheet_set_print_scale ws (round (100.0 * scale'))
  where
    scale' = min 0.1 (max 4.0 scale)

----------------------------------------------------------------------------------------------------

data LxwFormat_
newtype Format = Format { unFormat :: Ptr LxwFormat_ }

formatSetFontName :: Format -> String -> IO ()
formatSetFontName (Format fp) name =
  withCString name $ \cname ->
    format_set_font_name fp cname

formatSetFontSize :: Format -> Word16 -> IO ()
formatSetFontSize (Format fp) size =
  format_set_font_size fp size

data Color
  = ColorBlack
  | ColorBlue
  | ColorBrown
  | ColorCyan
  | ColorGray
  | ColorGreen
  | ColorLime
  | ColorMagenta
  | ColorNavy
  | ColorOrange
  | ColorPink
  | ColorPurple
  | ColorRed
  | ColorSilver
  | ColorWhite
  | ColorYellow
  | Color Word8 Word8 Word8

colorIndex :: Color -> Int32
colorIndex ColorBlack    = 0x00000000
colorIndex ColorBlue     = 0x000000ff
colorIndex ColorBrown    = 0x00800000
colorIndex ColorCyan     = 0x0000ffff
colorIndex ColorGray     = 0x00808080
colorIndex ColorGreen    = 0x00008000
colorIndex ColorLime     = 0x0000ff00
colorIndex ColorMagenta  = 0x00ff00ff
colorIndex ColorNavy     = 0x00000080
colorIndex ColorOrange   = 0x00ff6600
colorIndex ColorPink     = 0x00ff00ff
colorIndex ColorPurple   = 0x00800080
colorIndex ColorRed      = 0x00ff0000
colorIndex ColorSilver   = 0x00c0c0c0
colorIndex ColorWhite    = 0x00ffffff
colorIndex ColorYellow   = 0x00ffff00
colorIndex (Color r g b) =
  fromIntegral r `shiftL` 16 .|.
  fromIntegral g `shiftL`  8 .|.
  fromIntegral b

formatSetFontColor :: Format -> Color -> IO ()
formatSetFontColor (Format fp) color =
  format_set_font_color fp (colorIndex color)

formatSetNumFormat :: Format -> String -> IO ()
formatSetNumFormat (Format fp) fmt =
  withCString fmt $ \cfmt ->
    format_set_num_format fp cfmt

formatSetBold :: Format -> IO ()
formatSetBold (Format fp) =
  format_set_bold fp

formatSetItalic :: Format -> IO ()
formatSetItalic (Format fp) =
  format_set_italic fp

data UnderlineStyle
  = UnderlineNone
  | UnderlineSingle
  | UnderlineDouble
  | UnderlineSingleAccounting
  | UnderlineDoubleAccounting
  deriving (Eq, Enum, Read, Show)

formatSetUnderline :: Format -> UnderlineStyle -> IO ()
formatSetUnderline (Format fp) us =
  format_set_underline fp (fromIntegral (fromEnum us))

formatSetStrikeout :: Format -> IO ()
formatSetStrikeout (Format fp) =
  format_set_font_strikeout fp

data ScriptStyle
  = SuperScript
  | SubScript
  deriving (Eq, Enum, Read, Show)

formatSetScript :: Format -> ScriptStyle -> IO ()
formatSetScript (Format fp) s =
  format_set_font_script fp (1 + fromIntegral (fromEnum s))

formatSetBuiltInFormat :: Format -> Word8 -> IO ()
formatSetBuiltInFormat (Format fp) n =
  format_set_num_format_index fp n

data Align
  = AlignNone
  | AlignLeft
  | AlignCenter
  | AlignRight
  | AlignFill
  | AlignJustify
  | AlignCenterAcross
  | AlignDistributed
  deriving (Eq, Enum, Read, Show)

data VerticalAlign
  = VerticalAlignNone
  | VerticalAlignTop
  | VerticalAlignBottom
  | VerticalAlignCenter
  | VerticalAlignJustify
  | VerticalAlignDistributed
  deriving (Eq, Enum, Read, Show)

formatSetAlign :: Format -> Align -> IO ()
formatSetAlign (Format fp) a =
  format_set_align fp (fromIntegral (fromEnum a))

formatSetVerticalAlign :: Format -> VerticalAlign -> IO ()
formatSetVerticalAlign (Format fp) a =
  format_set_align fp a'
  where
    a' = case fromEnum a of
      0 -> 0
      n -> 7 + fromIntegral n

formatSetTextWrap :: Format -> IO ()
formatSetTextWrap (Format fp) =
  format_set_text_wrap fp

formatSetRotation :: Format -> Int -> IO ()
formatSetRotation (Format fp) angle =
  format_set_rotation fp (fromIntegral angle)

formatSetShrink :: Format -> IO ()
formatSetShrink (Format fp) =
  format_set_shrink fp

data Pattern
  = PatternNone
  | PatternSolid
  | PatternMediumGray
  | PatternDarkGray
  | PatternLightGray
  | PatternDarkHorizontal
  | PatternDarkVertical
  | PatternDarkDown
  | PatternDarkUp
  | PatternDarkGrid
  | PatternDarkTrellis
  | PatternLightHorizontal
  | PatternLightVertical
  | PatternLightDown
  | PatternLightUp
  | PatternLightGrid
  | PatternLightTrellis
  | PatternGray125
  | PatternGray0625
  deriving (Eq, Enum, Read, Show)

formatSetPattern :: Format -> Pattern -> IO ()
formatSetPattern (Format fp) pat =
  format_set_pattern fp (fromIntegral (fromEnum pat))

formatSetBackgroundColor :: Format -> Color -> IO ()
formatSetBackgroundColor (Format fp) color =
  format_set_bg_color fp (colorIndex color)

formatSetForegroundColor :: Format -> Color -> IO ()
formatSetForegroundColor (Format fp) color =
  format_set_fg_color fp (colorIndex color)

data Border
  = BorderAll
  | BorderBottom
  | BorderTop
  | BorderLeft
  | BorderRight
  deriving (Eq, Read, Show)

data BorderStyle
  = BorderNone
  | BorderThin
  | BorderMedium
  | BorderDashed
  | BorderDotted
  | BorderThick
  | BorderDouble
  | BorderHair
  | BorderMediumDashed
  | BorderDashDot
  | BorderMediumDashDot
  | BorderDashDotDot
  | BorderMediumDashDotDot
  | BorderSlantDashDot
  deriving (Eq, Enum, Read, Show)

formatSetBorder :: Format -> Border -> BorderStyle -> IO ()
formatSetBorder (Format fp) border style =
  function fp (fromIntegral (fromEnum style))
  where
    function = case border of
      BorderAll    -> format_set_border
      BorderBottom -> format_set_bottom
      BorderTop    -> format_set_top
      BorderLeft   -> format_set_left
      BorderRight  -> format_set_right

formatSetBorderColor :: Format -> Border -> Color -> IO ()
formatSetBorderColor (Format fp) border color =
  function fp (colorIndex color)
  where
    function = case border of
      BorderAll    -> format_set_border_color
      BorderBottom -> format_set_bottom_color
      BorderTop    -> format_set_top_color
      BorderLeft   -> format_set_left_color
      BorderRight  -> format_set_right_color

----------------------------------------------------------------------------------------------------

foreign import ccall "workbook_new"
  workbook_new :: CString -> IO (Ptr LxwWorkbook_)
foreign import ccall "workbook_new_opt"
  workbook_new_opt :: CString -> Ptr WorkbookOptions -> IO (Ptr LxwWorkbook_)
foreign import ccall "workbook_close"
  workbook_close :: Ptr LxwWorkbook_ -> IO ()
foreign import ccall "workbook_add_worksheet"
  workbook_add_worksheet :: Ptr LxwWorkbook_ -> CString ->
                            IO (Ptr LxwWorksheet_)
foreign import ccall "workbook_add_format"
  workbook_add_format :: Ptr LxwWorkbook_ -> IO (Ptr LxwFormat_)
foreign import ccall "workbook_set_properties"
  workbook_set_properties :: Ptr LxwWorkbook_ ->
                             Ptr DocProperties' -> IO ()
foreign import ccall "workbook_define_name"
  workbook_define_name :: Ptr LxwWorkbook_ -> CString -> CString -> IO ()
foreign import ccall "worksheet_write_number"
  worksheet_write_number :: Ptr LxwWorksheet_ ->
                            Word32 -> Word16 ->
                            Double -> Ptr LxwFormat_ -> IO ()
foreign import ccall "worksheet_write_string"
  worksheet_write_string :: Ptr LxwWorksheet_ ->
                            Word32 -> Word16 ->
                            CString -> Ptr LxwFormat_ -> IO ()
foreign import ccall "worksheet_write_formula"
  worksheet_write_formula :: Ptr LxwWorksheet_ ->
                             Word32 -> Word16 ->
                             CString -> Ptr LxwFormat_ -> IO ()
foreign import ccall "worksheet_write_array_formula"
  worksheet_write_array_formula :: Ptr LxwWorksheet_ ->
                                   Word32 -> Word16 ->
                                   Word32 -> Word16 ->
                                   CString -> Ptr LxwFormat_ -> IO ()
foreign import ccall "worksheet_write_datetime"
  worksheet_write_datetime :: Ptr LxwWorksheet_ ->
                              Word32 -> Word16 ->
                              Ptr DateTime -> Ptr LxwFormat_ -> IO ()
foreign import ccall "worksheet_write_url"
  worksheet_write_url :: Ptr LxwWorksheet_ ->
                         Word32 -> Word16 ->
                         CString -> Ptr LxwFormat_ -> IO ()
foreign import ccall "worksheet_set_row"
  worksheet_set_row :: Ptr LxwWorksheet_ ->
                       Word32 -> Double -> Ptr LxwFormat_ -> IO ()
foreign import ccall "worksheet_set_column"
  worksheet_set_column :: Ptr LxwWorksheet_ ->
                          Word16 -> Word16 -> Double -> Ptr LxwFormat_ -> IO ()
foreign import ccall "worksheet_insert_image"
  worksheet_insert_image :: Ptr LxwWorksheet_ ->
                            Word32 -> Word16 -> CString -> IO ()
foreign import ccall "worksheet_insert_image_opt"
  worksheet_insert_image_opt :: Ptr LxwWorksheet_ ->
                                Word32 -> Word16 ->
                                CString -> Ptr ImageOptions -> IO ()
foreign import ccall "worksheet_merge_range"
  worksheet_merge_range :: Ptr LxwWorksheet_ ->
                           Word32 -> Word16 ->
                           Word32 -> Word16 ->
                           CString -> Ptr LxwFormat_ -> IO ()
foreign import ccall "worksheet_freeze_panes"
  worksheet_freeze_panes :: Ptr LxwWorksheet_ ->
                            Word32 -> Word16 -> IO ()
foreign import ccall "worksheet_split_panes"
  worksheet_split_panes :: Ptr LxwWorksheet_ ->
                           Double -> Double -> IO ()
foreign import ccall "worksheet_set_landscape"
  worksheet_set_landscape :: Ptr LxwWorksheet_ -> IO ()
foreign import ccall "worksheet_set_portrait"
  worksheet_set_portrait :: Ptr LxwWorksheet_ -> IO ()
foreign import ccall "worksheet_set_page_view"
  worksheet_set_page_view :: Ptr LxwWorksheet_ -> IO ()
foreign import ccall "worksheet_set_paper"
  worksheet_set_paper :: Ptr LxwWorksheet_ -> Word8 -> IO ()
foreign import ccall "worksheet_set_margins"
  worksheet_set_margins :: Ptr LxwWorksheet_ ->
                           Double -> Double ->
                           Double -> Double -> IO ()
foreign import ccall "worksheet_set_header"
  worksheet_set_header :: Ptr LxwWorksheet_ -> CString -> IO ()
foreign import ccall "worksheet_set_footer"
  worksheet_set_footer :: Ptr LxwWorksheet_ -> CString -> IO ()
foreign import ccall "worksheet_set_zoom"
  worksheet_set_zoom :: Ptr LxwWorksheet_ -> Word16 -> IO ()
foreign import ccall "worksheet_set_print_scale"
  worksheet_set_print_scale:: Ptr LxwWorksheet_ -> Word16 -> IO ()

foreign import ccall "format_set_font_name"
  format_set_font_name :: Ptr LxwFormat_ -> CString -> IO ()
foreign import ccall "format_set_font_size"
  format_set_font_size :: Ptr LxwFormat_ -> Word16 -> IO ()
foreign import ccall "format_set_font_color"
  format_set_font_color :: Ptr LxwFormat_ -> Int32 -> IO ()
foreign import ccall "format_set_num_format"
  format_set_num_format :: Ptr LxwFormat_ -> CString -> IO ()
foreign import ccall "format_set_bold"
  format_set_bold :: Ptr LxwFormat_ -> IO ()
foreign import ccall "format_set_italic"
  format_set_italic :: Ptr LxwFormat_ -> IO ()
foreign import ccall "format_set_underline"
  format_set_underline :: Ptr LxwFormat_ -> Word8 -> IO ()
foreign import ccall "format_set_font_strikeout"
  format_set_font_strikeout :: Ptr LxwFormat_ -> IO ()
foreign import ccall "format_set_font_script"
  format_set_font_script :: Ptr LxwFormat_ -> Word8 -> IO ()
foreign import ccall "format_set_num_format_index"
  format_set_num_format_index :: Ptr LxwFormat_ -> Word8 -> IO ()
foreign import ccall "format_set_align"
  format_set_align :: Ptr LxwFormat_ -> Word8 -> IO ()
foreign import ccall "format_set_text_wrap"
  format_set_text_wrap :: Ptr LxwFormat_ -> IO ()
foreign import ccall "format_set_rotation"
  format_set_rotation :: Ptr LxwFormat_ -> Int16 -> IO ()
foreign import ccall "format_set_shrink"
  format_set_shrink :: Ptr LxwFormat_ -> IO ()
foreign import ccall "format_set_pattern"
  format_set_pattern :: Ptr LxwFormat_ -> Word8 -> IO ()
foreign import ccall "format_set_bg_color"
  format_set_bg_color :: Ptr LxwFormat_ -> Int32 -> IO ()
foreign import ccall "format_set_fg_color"
  format_set_fg_color :: Ptr LxwFormat_ -> Int32 -> IO ()
foreign import ccall "format_set_border"
  format_set_border :: Ptr LxwFormat_ -> Word8 -> IO ()
foreign import ccall "format_set_bottom"
  format_set_bottom :: Ptr LxwFormat_ -> Word8 -> IO ()
foreign import ccall "format_set_top"
  format_set_top :: Ptr LxwFormat_ -> Word8 -> IO ()
foreign import ccall "format_set_left"
  format_set_left :: Ptr LxwFormat_ -> Word8 -> IO ()
foreign import ccall "format_set_right"
  format_set_right :: Ptr LxwFormat_ -> Word8 -> IO ()
foreign import ccall "format_set_border_color"
  format_set_border_color :: Ptr LxwFormat_ -> Int32 -> IO ()
foreign import ccall "format_set_bottom_color"
  format_set_bottom_color :: Ptr LxwFormat_ -> Int32 -> IO ()
foreign import ccall "format_set_top_color"
  format_set_top_color :: Ptr LxwFormat_ -> Int32 -> IO ()
foreign import ccall "format_set_left_color"
  format_set_left_color :: Ptr LxwFormat_ -> Int32 -> IO ()
foreign import ccall "format_set_right_color"
  format_set_right_color :: Ptr LxwFormat_ -> Int32 -> IO ()
