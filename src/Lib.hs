{-# LANGUAGE DeriveGeneric             #-}
{-# LANGUAGE FlexibleContexts          #-}
{-# LANGUAGE NoMonomorphismRestriction #-}
{-# LANGUAGE OverloadedStrings         #-}
{-# LANGUAGE RecordWildCards           #-}
{-# LANGUAGE FlexibleInstances         #-}

module Lib (main, readSheet) where

import           Codec.Xlsx
import           Codec.Xlsx.Formatted
import           Control.Applicative        ((<|>))
import           Control.Exception          (throw)
import           Control.Lens
import           Control.Monad.IO.Class     (liftIO)
import           Control.Monad.State        (evalState, put)
import           Control.Monad.Trans.Except (ExceptT, runExceptT, throwE)
import qualified Data.Aeson                 as JSON
import qualified Data.ByteString            as BS
import qualified Data.ByteString.Lazy       as L
import qualified Data.List.NonEmpty         as NE
import qualified Data.Map.Strict            as M
import           Data.Maybe                 (fromMaybe)
import           Data.Monoid                ((<>))
import           Data.Text                  (Text)
import qualified Data.Text                  as T
import qualified Data.Yaml                  as YAML
import           Options.Generic            (getRecord)
import           System.Exit                (exitFailure)
import Xlsx2Yaml.CLI

-- | Extract a YAML file from an XLSX thingo
-- The first row of a sheet should have field names
-- data rows must have a key in the first column
type Row k v = M.Map k v
type Sheet r c v = M.Map (r, c) v

main :: IO ()
main = do
  Xlsx2YamlOpts {..} <- getRecord "xlsx2yaml - extract data from xlsx into yaml"
  -- Parameters
  let inFile = xlsx
      outFile = output
      dataStartRow = fromMaybe 4 beginDataRow
  --
  (xlsxx, styles) <- readXlsxFile inFile
  r <- runExceptT $ traverse (readSheet' dataStartRow styles xlsxx) sheet
  case r of
    Right done -> BS.writeFile outFile (YAML.encode (JSON.object (NE.toList done)))
    Left v -> do
      putStrLn $ "failed to find sheet " <> show v
      exitFailure

readXlsxFile :: FilePath -> IO (Xlsx, StyleSheet)
readXlsxFile inFile = do
  xlsxx <- toXlsx <$> L.readFile inFile
  styles <-
    case parseStyleSheet (_xlStyles xlsxx) of
      Left e -> throw e
      Right v -> pure v
  pure (xlsxx, styles)

-- used in testing
readSheet :: Int -> FilePath -> Text -> ExceptT T.Text IO (T.Text, YAML.Value)
readSheet dataStartRow inFile sheetToExtract =
     do (ss, styles) <- liftIO $ readXlsxFile inFile
        readSheet' dataStartRow styles ss sheetToExtract

readSheet'
  :: Monad m
  => Int -> StyleSheet -> Xlsx -> Text -> ExceptT Text m (Text, YAML.Value)
readSheet' dataStartRow styles ss sheetToExtract =
  let mkValue sheet =
        let sheetValue =
              sheetToValue dataStartRow styles sheet
        in (sheetToExtract, sheetValue)
  in case ss ^? ixSheet sheetToExtract of
       Just x -> pure (mkValue x)
       Nothing -> throwE sheetToExtract

-- | Encode a worksheet as a JSON Object.
-- First row is fields. Data rows start at R. Stop when you encounter Y blank rows
sheetToValue
  :: Int -> StyleSheet -> Worksheet -> JSON.Value
sheetToValue dataStart styles sheet =
  let fields = readFieldNames . extractRow 1 . _wsCells $ sheet
      dataRows = extractDefinedRows dataStart goodCells
      goodCells = definedCellValues styles sheet
  in JSON.toJSON $ fmap (JSON.object . mailMerge fields) dataRows

-- | Merge field names with a row of cells, into the contents of a JSON Object
mailMerge
  :: Ord k1
  => Row k1 k -> Row k1 OkCell -> [(k, JSON.Value)]
mailMerge fields =
  M.elems .
  M.intersectionWith f fields
  where f fieldName value = (fieldName, okCellToValue value)

readFieldNames :: Row k Cell -> Row k Text
readFieldNames cellsX =
  M.mapMaybe (^? cellValue . _Just . _CellText) cellsX

extractRow :: (Ord k2, Eq a1) => a1 -> Sheet a1 k2 a -> Row k2 a
-- extractRow :: (Ord k2, Eq a1) => a1 -> M.Map (a1, k2) a -> M.Map k2 a
extractRow selectedRow =
  M.mapKeys snd . M.filterWithKey (\rc _ -> fst rc == selectedRow)

-- Defined Rows are rows which have a value in the first column
-- Once enough (10) rows that are not defined have been found, we
-- stop searching
extractDefinedRows
  :: (Ord c, Num r, Num c, Eq r)
  => r -> M.Map (r, c) x -> [M.Map c x]
extractDefinedRows startRow sheet =
  evalState (action startRow) 0
  where
    stopWhenBadsIs = 100 :: Int -- will stop extracting rows when we encounter this many blanks in a row
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
          put 0 -- reset the empty row counter
          rest <- action (rowNum + 1)
          pure (row : rest)

data OkCell =
  OkCellValue CellValue | OkFilled Fill

definedCellValues :: StyleSheet -> Worksheet -> M.Map (Int, Int) OkCell
definedCellValues ss ws =
  let formattedCells' = toFormattedCells (_wsCells ws) [] ss
      p cell =
         (OkCellValue <$> _formattedValue cell) <|> (OkFilled <$> _formattedFill cell)

   in M.mapMaybe p formattedCells'

-- | Rendering to YAML/JSON

okCellToValue :: OkCell -> YAML.Value
okCellToValue (OkCellValue v) = cellValueToValue v
okCellToValue (OkFilled fill) = JSON.String (fillToValue fill)

-- | Fill as a string, if any of these accessors are Nothing, we get an empty string
fillToValue :: Fill -> Text
fillToValue f =
  -- Assume: For solid cell fills (no pattern), fgColor is used
  f ^. fillPattern . _Just . fillPatternFgColor . _Just . colorARGB . _Just

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
