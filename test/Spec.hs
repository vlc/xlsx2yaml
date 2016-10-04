{-# LANGUAGE OverloadedStrings #-}

import qualified Lib            as Lib

main :: IO ()
main = do
  Just _ <-
    Lib.readSheet 4 10 "test/sample-inputs/Design Specification.xlsx" "MASTER"
  pure ()
