{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE NoMonomorphismRestriction #-}

import Control.Monad.State
import qualified Data.Map.Strict as M
import qualified Data.Text as T
import Codec.Xlsx
import qualified Data.ByteString.Lazy as L
import Control.Lens
import Control.Lens.Extras (is)


-- | Extract a YAML file from an XLSX thingo

-- The first row of a sheet should have field names, in consecutive columns, no gaps
-- data rows start at row 4 (not sure they have to). they must at least have a key in the first column

main = do
  ss <- toXlsx <$> L.readFile "Design Specification.xlsx"
  let Just sheet = ss ^? ixSheet "MASTER"

  print (readFieldNames sheet)
  print (Prelude.length (extractDefinedRows 4 (view wsCells sheet)))

readFieldNames :: Worksheet -> [(Int, T.Text)]
readFieldNames ws =
  let cells = M.mapMaybe (^? _CellText) (definedCellValues ws)
      topRow = extractRow 1 cells
  in fmap fst . takeWhile (\((c, _), idx) -> c == idx) $
     zip (M.toList topRow) [1 ..]

extractRow :: (Ord k2, Eq a1) => a1 -> Map (a1, k2) a -> Map k2 a
extractRow selectedRow =
  M.mapKeys snd . M.filterWithKey (\rc _ -> fst rc == selectedRow)

-- Defined Rows
extractDefinedRows
  :: (Ord k2, Num a, Num k2, Eq a) =>
     a -> M.Map (a, k2) Cell -> [M.Map k2 Cell]
extractDefinedRows start sheet = evalState (action start) 0
  where
    action rowNum = do
      let row = extractRow rowNum sheet
          firstCell = M.lookup 1 row >>= view cellValue
      case firstCell of
        Nothing -> do
          modify succ
          numBads <- get
          if numBads > 10
            then pure []
            else action (rowNum + 1)
        Just _ -> do
          rest <- action (rowNum + 1)
          pure (row : rest)

definedCellValues :: Worksheet -> M.Map (Int, Int) CellValue
definedCellValues =
  M.mapMaybe id . fmap _cellValue . _wsCells

-- | Supporting Things

_CellText = prism' CellText g
 where g (CellText t) = Just t
       g _ = Nothing

_CellDouble  = prism' CellDouble g
 where g (CellDouble t) = Just t
       g _ = Nothing

_CellBool = prism' CellBool g
 where g (CellBool t) = Just t
       g _ = Nothing

_CellRich = prism' CellRich g
 where g (CellRich t) = Just t
       g _ = Nothing

