{-# LANGUAGE DeriveGeneric             #-}
{-# LANGUAGE NoMonomorphismRestriction #-}
{-# LANGUAGE OverloadedStrings         #-}
{-# LANGUAGE RecordWildCards           #-}

module Lib (main, readSheet) where

import           Codec.Xlsx
import           Control.Lens
import           Control.Monad.State
import qualified Data.Aeson           as JSON
import qualified Data.ByteString      as BS
import qualified Data.ByteString.Lazy as L
import qualified Data.Map.Strict      as M
import           Data.Maybe           (fromMaybe)
import           Data.Monoid          ((<>))
import qualified Data.Text            as T
import qualified Data.Yaml            as YAML
import           Options.Generic
import           System.Exit

-- | Extract a YAML file from an XLSX thingo
-- The first row of a sheet should have field names
-- data rows must have a key in the first column

data Opts = Opts
  { xlsx :: FilePath
  , sheet :: T.Text
  , output :: FilePath
  , beginDataRow :: Maybe Int
  , blanksBeforeStopping :: Maybe Int
  } deriving (Eq, Show, Generic)
instance ParseRecord Opts

main :: IO ()
main = do
  Opts {..} <- getRecord "xlsx2yaml - extract data from xlsx into yaml"
  -- Parameters
  let inFile = xlsx
      sheetToExtract = sheet
      outFile = output
      dataStartRow = fromMaybe 4 beginDataRow
      numSkipsBeforeStopping = fromMaybe 10 blanksBeforeStopping
  --
  r <- readSheet numSkipsBeforeStopping dataStartRow inFile sheetToExtract
  case r of
    Just done -> BS.writeFile outFile (YAML.encode done)
    Nothing -> do
      putStrLn $ "failed to find sheet " <> show sheetToExtract
      exitFailure

readSheet :: (Num t, Enum t, Ord t) => t -> Int -> FilePath -> Text -> IO (Maybe YAML.Value)
readSheet numSkipsBeforeStopping dataStartRow inFile sheetToExtract =
  let mkValue sheet =
        let sheetValue = sheetToValue numSkipsBeforeStopping dataStartRow sheet
        in JSON.object [(sheetToExtract, sheetValue)]
  in do ss <- toXlsx <$> L.readFile inFile
        pure $ fmap mkValue (ss ^? ixSheet sheetToExtract)

-- | Encode a worksheet as a JSON Object.
-- First row is fields. Data rows start at R. Stop when you encounter Y blank rows
sheetToValue
  :: (Ord t, Num t, Enum t)
  => t -> Int -> Worksheet -> JSON.Value
sheetToValue badsUntil dataStart sheet =
  let fields = readFieldNames goodCells
      dataRows = extractDefinedRows badsUntil dataStart goodCells
      goodCells = definedCellValues sheet
  in JSON.toJSON $ fmap (JSON.object . mailMerge fields) dataRows

mailMerge
  :: (Ord k1)
  => M.Map k1 k -> M.Map k1 CellValue -> [(k, JSON.Value)]
mailMerge fields =
  M.elems .
  M.intersectionWith f fields
  where f fieldName value = (fieldName, cellValueToValue value)

readFieldNames
  :: (Ord c, Num r, Eq r)
  => M.Map (r, c) CellValue -> M.Map c T.Text
readFieldNames cellsX =
  let cells = M.mapMaybe (^? _CellText) cellsX
   in extractRow 1 cells

extractRow :: (Ord k2, Eq a1) => a1 -> M.Map (a1, k2) a -> M.Map k2 a
extractRow selectedRow =
  M.mapKeys snd . M.filterWithKey (\rc _ -> fst rc == selectedRow)

-- Defined Rows are rows which have a value in the first column
-- Once enough (10) rows that are not defined have been found, we
-- stop searching
extractDefinedRows
  :: (Ord c, Num r, Num c, Eq r, Enum z, Ord z, Num z)
  => z -> r -> M.Map (r, c) CellValue -> [M.Map c CellValue]
extractDefinedRows stopWhenBadsIs startRow sheet = evalState (action startRow) 0
  where
    action rowNum = do
      let row = extractRow rowNum sheet
          firstCell = row ^? ix 1
      case firstCell of
        Nothing -> do
          numBads <- id <%= succ
          if numBads > stopWhenBadsIs
            then pure []
            else action (rowNum + 1)
        Just _ -> do
          rest <- action (rowNum + 1)
          pure (row : rest)

definedCellValues :: Worksheet -> M.Map (Int, Int) CellValue
definedCellValues =
  M.mapMaybe _cellValue . _wsCells

cellValueToValue :: CellValue -> JSON.Value
cellValueToValue (CellText t) = JSON.String t
cellValueToValue (CellDouble d) = JSON.Number (fromRational . toRational $ d)
cellValueToValue (CellBool b) = JSON.Bool b
cellValueToValue (CellRich b) = JSON.String (b ^. traverse . richTextRunText)

-- | Supporting Things

_CellText :: Prism' CellValue T.Text
_CellText = prism' CellText g
 where g (CellText t) = Just t
       g _ = Nothing

_CellDouble :: Prism' CellValue Double
_CellDouble  = prism' CellDouble g
 where g (CellDouble t) = Just t
       g _ = Nothing

_CellBool :: Prism' CellValue Bool
_CellBool = prism' CellBool g
 where g (CellBool t) = Just t
       g _ = Nothing

_CellRich :: Prism' CellValue [RichTextRun]
_CellRich = prism' CellRich g
 where g (CellRich t) = Just t
       g _ = Nothing
